using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Data;
using System.Runtime.InteropServices;


namespace BDO_Localisation_AddOn
{
    static partial class BDO_TaxInvoiceReceived
    {
        public static void createDocumentUDO(out string errorText)
        {
            errorText = null;
            string tableName = "BDO_TAXR";
            string description = "Tax Invoice Received";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_Document, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;
            List<string> listValidValues;

            fieldskeysMap = new Dictionary<string, object>(); //ბიზნესპარტნიორი (კოდი)
            fieldskeysMap.Add("Name", "cardCode");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Supplier Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ბიზნესპარტნიორი (სახელი)
            fieldskeysMap.Add("Name", "cardCodeN");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Supplier Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ბიზნესპარტნიორი (გსნ)
            fieldskeysMap.Add("Name", "cardCodeT");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Supplier TIN");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //მიღების თარიღი
            fieldskeysMap.Add("Name", "recvDate");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Receive Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ID
            fieldskeysMap.Add("Name", "invID");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Invoice ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ნომერი
            fieldskeysMap.Add("Name", "number");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Invoice Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //სერია
            fieldskeysMap.Add("Name", "series");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Invoice Series");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>(); //სტატუსი
            listValidValuesDict.Add("empty", "");
            listValidValuesDict.Add("paper", "Paper"); //ქაღალდის
            listValidValuesDict.Add("received", "Received"); //მიღებული
            listValidValuesDict.Add("confirmed", "Confirmed"); //დადასტურებული
            listValidValuesDict.Add("incompleteReceived", "Incomplete Received"); //არასრულად მიღებული
            listValidValuesDict.Add("denied", "Denied"); //უარყოფილი
            listValidValuesDict.Add("cancellationProcess", "Cancellation Process"); //გაუქმების პროცესში
            listValidValuesDict.Add("canceled", "Canceled"); //გაუქმებული
            listValidValuesDict.Add("correctionReceived", "Correction Received"); //მიღებული კორექტირებული
            listValidValuesDict.Add("correctionDenied", "Correction Denied"); //უარყოფილი კორექტირებული
            listValidValuesDict.Add("correctionConfirmed", "Correction Confirmed"); //დადასტურებული კორექტირებული
            listValidValuesDict.Add("attachedToTheDeclaration", "Attached To The Declaration"); //დეკლარაციაზე მიბმული
            listValidValuesDict.Add("removed", "Removed"); //წაშლილი
            listValidValuesDict.Add("corrected", "Corrected"); //კორექტირებული
            listValidValuesDict.Add("replaced", "Replaced"); //ჩანაცვლებული

            fieldskeysMap = new Dictionary<string, object>(); //სტატუსი
            fieldskeysMap.Add("Name", "status");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Invoice Status");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დამდასტურებელი პირი (ID)
            fieldskeysMap.Add("Name", "confInfo");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Confirmation Info");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დამდასტურებელი პირი (სახელი)
            fieldskeysMap.Add("Name", "confInfN");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Confirmation Info Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დადასტურების თარიღი
            fieldskeysMap.Add("Name", "confDate");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Confirmation Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დეკლარაციის თვე
            fieldskeysMap.Add("Name", "declDate");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Declaration Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დეკლარაციის ნომერი
            fieldskeysMap.Add("Name", "declNumber");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Declaration Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ოპერაციის თვე
            fieldskeysMap.Add("Name", "opDate");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Operation Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amount"); //თანხა დღგ-ის ჩათვლით
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amountTX"); //დღგ-ის თანხა
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Vat Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //მაკორექტირებელი ანგარიშ–ფაქტურისთვის
            fieldskeysMap.Add("Name", "corrInv");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "For Correction");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtCor"); //თანხა დღგ-ის ჩათვლით (დაკორექტირებული ფაქტურის)
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Amount Correction");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtTXCor"); //დღგ-ის თანხა (დაკორექტირებული ფაქტურის)
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Tax Amount Correction");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtACor"); //თანხა დღგ-ის ჩათვლით (კორექტირების შემდეგ)
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Amount After Correction");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtTXACr"); //დღგ-ის თანხა (კორექტირების შემდეგ)
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Tax Amount After Correction");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            listValidValues = new List<string>(); //კორექტირების მიზეზები
            listValidValues.Add(""); //-1
            listValidValues.Add("Canceled Tax Operation"); //1 //გაუქმებულია დასაბეგრი ოპერაცია
            listValidValues.Add("Changed Tax Operation Type"); //2 //შეცვლილია დასაბეგრი ოპერაციის სახე
            listValidValues.Add("Changed Agreement Amount Prices Decrease"); //3 //ფასების შემცირების ან სხვა მიზეზით შეცვლილია ოპერაციაზე ადრე შეთანხმებული კომპენსაციის თანხა
            listValidValues.Add("Item Service Returned To Seller"); //4 საქონელი (მომსახურება) სრულად ან ნაწილობრივ უბრუნდება გამყიდველს

            fieldskeysMap = new Dictionary<string, object>(); //კორექტირების მიზეზები
            fieldskeysMap.Add("Name", "corrType");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Correction Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ელექტრონული სახით
            fieldskeysMap.Add("Name", "elctrnic");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Electronic");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "Y");

            UDO.addUserTableFields(fieldskeysMap, out errorText);


            //------------------------------
            fieldskeysMap = new Dictionary<string, object>(); //ჩათვლილი (კომბო)
            fieldskeysMap.Add("Name", "TaxInRcvd");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Tax Invoice Received");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //------------------------------


            //fieldskeysMap = new Dictionary<string, object>(); //ჩათვლილია
            //fieldskeysMap.Add("Name", "vatRecvd");
            //fieldskeysMap.Add("TableName", "BDO_TAXR");
            //fieldskeysMap.Add("Description", "Vat Received");
            //fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            //fieldskeysMap.Add("EditSize", 1);
            //fieldskeysMap.Add("DefaultValue", "N");

            //UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ჩათვლის თვე
            fieldskeysMap.Add("Name", "vatRDate");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Vat Receive Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //კომენტარი
            fieldskeysMap.Add("Name", "comment");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Comment");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "corrDoc"); //კორექტირების დოკუმენტი
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Correction Document");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "corrDTxt"); //კორექტირების დოკუმენტი (TXT)
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Correction Document");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "corrDocID"); //კორექტირების დოკუმენტის ID
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Correction Document ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            listValidValues = new List<string>();
            listValidValues.Add("Not Linked"); //0 //არ არის მიბმული
            listValidValues.Add("Linked"); //1 //მიბმულია
            listValidValues.Add("Linked Partial"); //2 //ნაწილობრივ მიბმული

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "LinkStatus"); //მიბმის სტატუსი
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Link Status");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ავანსის ანგარიშ–ფაქტურა
            fieldskeysMap.Add("Name", " downPaymnt");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Down Payment");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // DocDate 
            fieldskeysMap.Add("Name", "docDate");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Posting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // JrnEnt 
            fieldskeysMap.Add("Name", "JrnEnt");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Journal Entry");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დღგ-ის თარიღი
            fieldskeysMap.Add("Name", "taxDate");
            fieldskeysMap.Add("TableName", "BDO_TAXR");
            fieldskeysMap.Add("Description", "Tax Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            tableName = "BDO_TXR1";
            description = "Tax Invoice Received Child1";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            if (result != 0)
            {
                return;
            }

            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add("APInvoice"); //0
            listValidValues.Add("APCreditNote"); //1
            listValidValues.Add("APDownPaymentInvoice"); //2

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "baseDocT"); //საფუძველი დოკუმენტის ტიპი
            fieldskeysMap.Add("TableName", "BDO_TXR1");
            fieldskeysMap.Add("Description", "Base Document Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            bool result2 = UDO.addNewValidValuesUserFieldsMD("@BDO_TXR1", "baseDocT", "2", "APDownPaymentInvoice", out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "baseDoc"); //საფუძველი დოკუმენტი
            fieldskeysMap.Add("TableName", "BDO_TXR1");
            fieldskeysMap.Add("Description", "Base Document");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "baseDTxt"); //საფუძველი დოკუმენტი (TXT)
            fieldskeysMap.Add("TableName", "BDO_TXR1");
            fieldskeysMap.Add("Description", "Base Document");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtBsDc"); //საფუძველი დოკუმენტის თანხა
            fieldskeysMap.Add("TableName", "BDO_TXR1");
            fieldskeysMap.Add("Description", "Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "tAmtBsDc"); //საფუძველი დოკუმენტის დღგ-ის თანხა
            fieldskeysMap.Add("TableName", "BDO_TXR1");
            fieldskeysMap.Add("Description", "Vat Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ზედნადების ნომერი
            fieldskeysMap.Add("Name", "wbNumber");
            fieldskeysMap.Add("TableName", "BDO_TXR1");
            fieldskeysMap.Add("Description", "Waybill Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            tableName = "BDO_TXR2";
            description = "Tax Invoice Received Child2";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            if (result != 0)
            {
                return;
            }

            fieldskeysMap = new Dictionary<string, object>(); //ზედნადების ნომერი
            fieldskeysMap.Add("Name", "wbNumber");
            fieldskeysMap.Add("TableName", "BDO_TXR2");
            fieldskeysMap.Add("Description", "Waybill Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            tableName = "BDO_TXR3";
            description = "Tax Invoice Received Child3";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            if (result != 0)
            {
                return;
            }

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "g_number"); //რაოდენობა
            fieldskeysMap.Add("TableName", "BDO_TXR3");
            fieldskeysMap.Add("Description", "Quantity");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "full_amount"); //თანხა დღგ-ის და აქციზის ჩათვლით
            fieldskeysMap.Add("TableName", "BDO_TXR3");
            fieldskeysMap.Add("Description", "full amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "drg_amount"); //დღგ
            fieldskeysMap.Add("TableName", "BDO_TXR3");
            fieldskeysMap.Add("Description", "Vat Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "goods");
            fieldskeysMap.Add("TableName", "BDO_TXR3");
            fieldskeysMap.Add("Description", "Item");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 250);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "g_unit");
            fieldskeysMap.Add("TableName", "BDO_TXR3");
            fieldskeysMap.Add("Description", "Unit");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 250);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "vat_type");
            fieldskeysMap.Add("TableName", "BDO_TXR3");
            fieldskeysMap.Add("Description", "Vat type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 250);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            tableName = "BDO_TXR4";
            description = "Tax Invoice Received Child4";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            if (result != 0)
            {
                return;
            }

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "max_amount"); //დღგ მაქს
            fieldskeysMap.Add("TableName", "BDO_TXR4");
            fieldskeysMap.Add("Description", "max vat Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "drg_amount"); //დღგ
            fieldskeysMap.Add("TableName", "BDO_TXR4");
            fieldskeysMap.Add("Description", "Vat Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DP_invoice");
            fieldskeysMap.Add("TableName", "BDO_TXR4");
            fieldskeysMap.Add("Description", "Down Payment Invoice");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 250);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            tableName = "BDO_TXR5";
            description = "Tax Invoice Received Child5";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            if (result != 0)
            {
                return;
            }

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "tax_invoice");
            fieldskeysMap.Add("TableName", "BDO_TXR5");
            fieldskeysMap.Add("Description", "Tax Invoice");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 250);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "drg_amount"); //დღგ
            fieldskeysMap.Add("TableName", "BDO_TXR5");
            fieldskeysMap.Add("Description", "Vat Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void registerUDO(out string errorText)
        {
            errorText = null;
            string code = "UDO_F_BDO_TAXR_D"; //20 characters (must include at least one alphabetical character).
            Dictionary<string, object> formProperties;

            formProperties = new Dictionary<string, object>();
            formProperties.Add("Name", "Tax Invoice Received"); //100 characters
            formProperties.Add("TableName", "BDO_TAXR");
            formProperties.Add("ObjectType", SAPbobsCOM.BoUDOObjType.boud_Document);
            formProperties.Add("CanCancel", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanClose", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanCreateDefaultForm", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanDelete", SAPbobsCOM.BoYesNoEnum.tNO);
            formProperties.Add("CanFind", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanYearTransfer", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("ManageSeries", SAPbobsCOM.BoYesNoEnum.tYES);

            List<Dictionary<string, object>> listFindColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listFormColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listChildTables = new List<Dictionary<string, object>>();

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_cardCode");
            fieldskeysMap.Add("ColumnDescription", "Supplier Code"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_cardCodeN");
            fieldskeysMap.Add("ColumnDescription", "Supplier Name"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_cardCodeT");
            fieldskeysMap.Add("ColumnDescription", "Supplier TIN"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_series");
            fieldskeysMap.Add("ColumnDescription", "Invoice Series"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_number");
            fieldskeysMap.Add("ColumnDescription", "Invoice Number"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_invID");
            fieldskeysMap.Add("ColumnDescription", "Invoice ID"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_status");
            fieldskeysMap.Add("ColumnDescription", "Invoice Status"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_opDate");
            fieldskeysMap.Add("ColumnDescription", "Operation Date"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_amount");
            fieldskeysMap.Add("ColumnDescription", "Amount"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_amountTX");
            fieldskeysMap.Add("ColumnDescription", "Vat Amount"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_amtACor");
            fieldskeysMap.Add("ColumnDescription", "Amount After Correction"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_amtTXACr");
            fieldskeysMap.Add("ColumnDescription", "Tax Amount After Correction"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_corrDocID");
            fieldskeysMap.Add("ColumnDescription", "Correction Document ID"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "DocEntry");
            fieldskeysMap.Add("ColumnDescription", "DocEntry"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "DocNum");
            fieldskeysMap.Add("ColumnDescription", "Number"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "Status");
            fieldskeysMap.Add("ColumnDescription", "Status"); //30 characters
            listFindColumns.Add(fieldskeysMap);

            //----------------------------------
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_TaxInRcvd");
            fieldskeysMap.Add("ColumnDescription", "Tax Invoice Received"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            //----------------------------------


            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "CreateDate");
            fieldskeysMap.Add("ColumnDescription", "Create Date"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            //fieldskeysMap.Add("ColumnAlias", "UpdateDate");
            //fieldskeysMap.Add("ColumnDescription", "Update Date"); //30 characters

            fieldskeysMap.Add("ColumnAlias", "U_docDate");
            fieldskeysMap.Add("ColumnDescription", "Posting Date"); //30 characters

            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "Remark");
            fieldskeysMap.Add("ColumnDescription", "Remark"); //30 characters
            listFindColumns.Add(fieldskeysMap);

            formProperties.Add("FindColumns", listFindColumns);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("FormColumnAlias", "DocEntry");
            fieldskeysMap.Add("FormColumnDescription", "DocEntry"); //30 characters
            listFormColumns.Add(fieldskeysMap);

            formProperties.Add("FormColumns", listFormColumns);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("TableName", "BDO_TXR1");
            fieldskeysMap.Add("ObjectName", "BDO_TXR1"); //30 characters
            listChildTables.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("TableName", "BDO_TXR2");
            fieldskeysMap.Add("ObjectName", "BDO_TXR2"); //30 characters
            listChildTables.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("TableName", "BDO_TXR3");
            fieldskeysMap.Add("ObjectName", "BDO_TXR3"); //30 characters
            listChildTables.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("TableName", "BDO_TXR4");
            fieldskeysMap.Add("ObjectName", "BDO_TXR4"); //30 characters
            listChildTables.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("TableName", "BDO_TXR5");
            fieldskeysMap.Add("ObjectName", "BDO_TXR5"); //30 characters
            listChildTables.Add(fieldskeysMap);

            formProperties.Add("ChildTables", listChildTables);

            UDO.registerUDO(code, formProperties, out errorText);

            GC.Collect();
        }

        public static void updateUDO()
        {

            string code = "UDO_F_BDO_TAXR_D";

            List<string> listChildTables = new List<string>(); //ChildTables
            listChildTables.Add("BDO_TXR1");
            listChildTables.Add("BDO_TXR2");
            listChildTables.Add("BDO_TXR3");
            listChildTables.Add("BDO_TXR4");
            listChildTables.Add("BDO_TXR5");

            string queryChildTables = @"SELECT ""TableName""
                                        FROM ""UDO1"" 
                                        WHERE ""Code"" = '" + code + "'";

            for (int i = 0; i < listChildTables.Count(); i++)
            {
                string conTxt = (i == 0 ? " AND ( " : " OR ");
                queryChildTables = queryChildTables + conTxt + @" ""TableName"" = '" + listChildTables[i] + "'";
            }
            queryChildTables = queryChildTables + " )";

            List<string> listFindColumns = new List<string>(); //FindColumns
            listFindColumns.Add("U_cardCode");
            listFindColumns.Add("U_cardCodeN");
            listFindColumns.Add("U_cardCodeT");
            listFindColumns.Add("U_series");
            listFindColumns.Add("U_number");
            listFindColumns.Add("U_invID");
            listFindColumns.Add("U_status");
            listFindColumns.Add("U_opDate");
            listFindColumns.Add("U_amount");
            listFindColumns.Add("U_amountTX");
            listFindColumns.Add("U_amtACor");
            listFindColumns.Add("U_amtTXACr");
            listFindColumns.Add("U_corrDocID");
            listFindColumns.Add("DocEntry");
            listFindColumns.Add("DocNum");
            listFindColumns.Add("Status");
            listFindColumns.Add("U_TaxInRcvd");
            listFindColumns.Add("CreateDate");
            listFindColumns.Add("U_docDate");
            listFindColumns.Add("Remark");

            string queryFindColumns = @"SELECT ""ColAlias""
                                FROM ""UDO2"" 
                                WHERE ""Code"" = '" + code + "'";
            for (int i = 0; i < listFindColumns.Count(); i++)
            {
                string conTxt = (i == 0 ? " AND ( " : " OR ");
                queryFindColumns = queryFindColumns + conTxt + @" ""ColAlias"" = '" + listFindColumns[i] + "'";
            }
            queryFindColumns = queryFindColumns + " )";

            SAPbobsCOM.Recordset oRecordSetFindColumns = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSetFindColumns.DoQuery(queryFindColumns);

            SAPbobsCOM.Recordset oRecordSetChildTables = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSetChildTables.DoQuery(queryChildTables);

            if (oRecordSetFindColumns.RecordCount != listFindColumns.Count || oRecordSetChildTables.RecordCount != listChildTables.Count)
            {
                Marshal.ReleaseComObject(oRecordSetFindColumns);
                oRecordSetFindColumns = null;

                Marshal.ReleaseComObject(oRecordSetChildTables);
                oRecordSetChildTables = null;
                GC.WaitForPendingFinalizers();

                string errorText = "";
                registerUDO(out errorText);
            }
            else
            {
                Marshal.ReleaseComObject(oRecordSetFindColumns);
                oRecordSetFindColumns = null;

                Marshal.ReleaseComObject(oRecordSetChildTables);
                oRecordSetChildTables = null;
                GC.WaitForPendingFinalizers();
            }
        }

        public static void addMenus()
        {
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                fatherMenuItem = Program.uiApp.Menus.Item("2304");
                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "UDO_F_BDO_TAXR_D";
                oCreationPackage.String = BDOSResources.getTranslate("TaxInvoiceReceived");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            oForm.AutoManaged = true;

            errorText = null;
            Dictionary<string, object> formItems;
            string itemName = "";

            int left_s = 6;
            int left_e = 127;
            int height = 15;
            int top = 6;
            int width_s = 121;
            int width_e = 148;

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "cardCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("SupplierCode"));
            formItems.Add("LinkTo", "cardCodeE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            bool multiSelection = false;
            string objectType = "2"; //SAPbouiCOM.BoLinkedObject.lf_BusinessPartner, Business Partner object 
            string uniqueID_lf_BusinessPartnerCFL = "BusinessPartner_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_BusinessPartnerCFL);

            //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
            SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_BusinessPartnerCFL);
            SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
            SAPbouiCOM.Condition oCon = oCons.Add();
            oCon.Alias = "CardType";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "S"; //მომწოდებელი
            oCFL.SetConditions(oCons);

            formItems = new Dictionary<string, object>();
            itemName = "cardCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_cardCode");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_BusinessPartnerCFL);
            formItems.Add("ChooseFromListAlias", "CardCode");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "cardCodeNE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_cardCodeN");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 4);
            formItems.Add("Width", width_e * 3 / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("cardCodeNE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            formItems = new Dictionary<string, object>();
            itemName = "cardCodeLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "cardCodeE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "cardCodeTS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("SupplierTIN"));
            formItems.Add("LinkTo", "cardCodeTE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "cardCodeTE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_cardCodeT");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("cardCodeTE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "recvDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Received")); //მიღებულია
            formItems.Add("LinkTo", "recvDateE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "recvDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_recvDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "corrInvCH"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_corrInv");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e * 1.05);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Caption", BDOSResources.getTranslate("CorrectionTaxInvoice"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("corrInvCH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            formItems = new Dictionary<string, object>();
            itemName = "elctrnicCH"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_elctrnic");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Left", left_s + width_e + 10);
            formItems.Add("Width", width_e * 2 / 3 + 10);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Caption", BDOSResources.getTranslate("Electronic"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "downPmntCH"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_downPaymnt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e * 1.26);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Caption", BDOSResources.getTranslate("DownPaymentInvoice"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("downPmntCH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            top = top + height + 1;
            //top = top + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSJrnEnS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TransactionNo"));
            formItems.Add("LinkTo", "BDOSJrnEnt");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSJrnEnt";
            formItems.Add("Size", 20);
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_JrnEnt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Enabled", false);
            formItems.Add("AffectsFormMode", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("BDOSJrnEnt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            formItems = new Dictionary<string, object>();
            itemName = "BDOSJEntLB";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDOSJrnEnt");
            formItems.Add("LinkedObjectType", "30");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //კორექტირების დოკუმენტის ტიპის მიხედვით CFL- ის დამატება ---->           
            multiSelection = false;
            objectType = "UDO_F_BDO_TAXR_D"; //Tax Invoice Received
            string uniqueID_CorrDocCFL = "CorrDoc_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_CorrDocCFL);
            //<----

            left_s = 295;
            left_e = left_s + 121;

            formItems = new Dictionary<string, object>();
            itemName = "corrDocIDS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CorrectionDocumentID"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "corrDocIDE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_corrDocID");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("corrDocIDE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            formItems = new Dictionary<string, object>();
            itemName = "corrDocE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_corrDTxt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 2 + 20);
            formItems.Add("Width", width_e / 2 - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("corrDocE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            formItems = new Dictionary<string, object>();
            itemName = "corrDocE1"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_corrDoc");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 2 + 20);
            formItems.Add("Width", width_e / 2 - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "corrDocLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e + width_e / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "corrDocE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            List<string> listValidValues = new List<string>(); //კორექტირების მიზეზები
            listValidValues.Add(""); //-1
            listValidValues.Add(BDOSResources.getTranslate("CanceledTaxOperation")); //1 //გაუქმებულია დასაბეგრი ოპერაცია
            listValidValues.Add(BDOSResources.getTranslate("ChangedTaxOperationType")); //2 //შეცვლილია დასაბეგრი ოპერაციის სახე
            listValidValues.Add(BDOSResources.getTranslate("ChangedAgreementAmountPricesDecrease")); //3 //ფასების შემცირების ან სხვა მიზეზით შეცვლილია ოპერაციაზე ადრე შეთანხმებული კომპენსაციის თანხა
            listValidValues.Add(BDOSResources.getTranslate("ItemServiceReturnedToSeller")); //4 საქონელი (მომსახურება) სრულად ან ნაწილობრივ უბრუნდება გამყიდველს

            formItems = new Dictionary<string, object>();
            itemName = "corrTypeCB"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_corrType");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e + width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Description", BDOSResources.getTranslate("Reason"));
            formItems.Add("Enabled", false);
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("corrTypeCB").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            //left_s = 295;
            //left_e = left_s + 121;
            top = 6;

            formItems = new Dictionary<string, object>();
            itemName = "No.S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s / 3 + 3);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Number"));
            formItems.Add("LinkTo", "SeriesC");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "SeriesC"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "Series");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_s + width_s / 3 + 3);
            formItems.Add("Width", width_s * 2 / 3 - 3);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Description", BDOSResources.getTranslate("Series"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "DocNumE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "DocNum");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "StatusS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Status"));
            formItems.Add("LinkTo", "StatusC");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "StatusC"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "Status");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Description", BDOSResources.getTranslate("Status"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "CreateDatS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CreateDate"));
            formItems.Add("LinkTo", "CreateDatE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "CreateDatE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "CreateDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("CreateDatE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "DocDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("PostingDate"));
            formItems.Add("LinkTo", "DocDateE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "DocDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_docDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("DocDateE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "taxDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TaxDate"));
            formItems.Add("LinkTo", "taxDateE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "taxDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_taxDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Description", "(" + BDOSResources.getTranslate("DownPaymentInvoice") + ")");
            //formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            //oForm.Items.Item("taxDateE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            left_s = 6;
            left_e = 127;

            //ელექტრონული ანგარიშ-ფაქტურა ---->
            //top = top + height + 1;
            top = 86 + 40;

            formItems = new Dictionary<string, object>();
            itemName = "taxInvS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s - 5);
            formItems.Add("Width", width_s * 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ElectricTaxInvoice"));
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 12);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 25;

            formItems = new Dictionary<string, object>();
            itemName = "seriesS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("SeriesNumber"));
            formItems.Add("LinkTo", "seriesE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "seriesE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_series");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("seriesE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            formItems = new Dictionary<string, object>();
            itemName = "numberE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_number");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 2);
            formItems.Add("Width", width_e / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("numberE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "statusS1"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TaxInvoceStatus"));
            formItems.Add("LinkTo", "statusCB");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>(); //სტატუსი
            listValidValuesDict.Add("empty", "");
            listValidValuesDict.Add("paper", BDOSResources.getTranslate("Paper")); //ქაღალდის
            listValidValuesDict.Add("received", BDOSResources.getTranslate("Received")); //მიღებული
            listValidValuesDict.Add("confirmed", BDOSResources.getTranslate("Confirmed")); //დადასტურებული
            listValidValuesDict.Add("incompleteReceived", BDOSResources.getTranslate("IncompleteReceived")); //არასრულად მიღებული
            listValidValuesDict.Add("denied", BDOSResources.getTranslate("Denied")); //უარყოფილი
            listValidValuesDict.Add("cancellationProcess", BDOSResources.getTranslate("CancellationProcess")); //გაუქმების პროცესში
            listValidValuesDict.Add("canceled", BDOSResources.getTranslate("Canceled")); //გაუქმებული
            listValidValuesDict.Add("correctionReceived", BDOSResources.getTranslate("CorrectionReceived")); //მიღებული კორექტირებული
            listValidValuesDict.Add("correctionDenied", BDOSResources.getTranslate("CorrectionDenied")); //უარყოფილი კორექტირებული
            listValidValuesDict.Add("correctionConfirmed", BDOSResources.getTranslate("CorrectionConfirmed")); //დადასტურებული კორექტირებული
            listValidValuesDict.Add("attachedToTheDeclaration", BDOSResources.getTranslate("AttachedToTheDeclaration")); //დეკლარაციაზე მიბმული
            listValidValuesDict.Add("removed", BDOSResources.getTranslate("Removed")); //წაშლილი
            listValidValuesDict.Add("corrected", BDOSResources.getTranslate("Corrected")); //კორექტირებული
            listValidValuesDict.Add("replaced", BDOSResources.getTranslate("Replaced")); //ჩანაცვლებული

            formItems = new Dictionary<string, object>();
            itemName = "statusCB"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_status");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Description", BDOSResources.getTranslate("TaxInvoceStatus"));
            formItems.Add("Enabled", false);
            formItems.Add("ValidValues", listValidValuesDict);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("statusCB").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "confInfoS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Approver"));
            formItems.Add("LinkTo", "confInfoE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            multiSelection = false;
            objectType = "12"; //SAPbouiCOM.BoLinkedObject.lf_User
            string uniqueID_lf_UsersCFL = "Users_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_UsersCFL);

            formItems = new Dictionary<string, object>();
            itemName = "confInfoE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_confInfo");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_UsersCFL);
            formItems.Add("ChooseFromListAlias", "USERID");
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("confInfoE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            formItems = new Dictionary<string, object>();
            itemName = "confInfoE1"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_confInfN");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 4);
            formItems.Add("Width", width_e * 3 / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("confInfoE1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            formItems = new Dictionary<string, object>();
            itemName = "confInfoLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "confInfoE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "confDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ConfirmationDate"));
            formItems.Add("LinkTo", "confDateE");


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "confDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_confDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("confDateE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "declNumS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s * 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Declaration"));
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 12);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "declDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DeclarationDate"));
            formItems.Add("LinkTo", "declDateE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "declDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_declDate");
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
            //oForm.Items.Item("declDateE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "declNmberS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DeclarationNumber"));
            formItems.Add("LinkTo", "declNmberE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "declNmberE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_declNumber");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("declNmberE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            left_s = 295;
            left_e = left_s + 121;
            top = 86 + 40;
            top = top + 25;

            formItems = new Dictionary<string, object>();
            itemName = "invIDS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TaxInvoiceID"));
            formItems.Add("LinkTo", "invIDE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "invIDE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_invID");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("invIDE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "opDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("OperationMonth"));
            formItems.Add("LinkTo", "opDateE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "opDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_opDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("opDateE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "amountS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("AmountWithVat"));
            formItems.Add("LinkTo", "amountE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "amountE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_amount");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("amountE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "amountTXS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("VatAmount"));
            formItems.Add("LinkTo", "amountTXE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "amountTXE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_amountTX");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("amountTXE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            top = 2 * (top + height + 1);

            formItems = new Dictionary<string, object>();
            itemName = "declStatS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("AttachedToTheDeclaration"));
            formItems.Add("LinkTo", "declStatS1");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "declStatS1"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("NotLinked"));
            formItems.Add("TextStyle", 4);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            left_s = 6;
            left_e = 127;

            //კორექტირებული ფაქტურის მონაცემები ---->

            top = 223 + 40;

            formItems = new Dictionary<string, object>();
            itemName = "21_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s * 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CorrectedTaxInvoiceData"));
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 12);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 25;

            formItems = new Dictionary<string, object>();
            itemName = "amtCorS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("AmountWithVat")); //(კორექტირებული)
            formItems.Add("LinkTo", "amtCorE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "amtCorE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_amtCor");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("amtCorE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "amtTXCorS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("VatAmount")); //(კორექტირებული)
            formItems.Add("LinkTo", "amtTXCorE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "amtTXCorE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_amtTXCor");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("amtTXCorE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            left_s = 295;
            left_e = left_s + 121;

            //მონაცემები კორექტირების შემდეგ ---->

            top = 223 + 40;

            formItems = new Dictionary<string, object>();
            itemName = "24_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s * 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DataAfterCorrection"));
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 12);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 25;

            formItems = new Dictionary<string, object>();
            itemName = "amtACorS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("AmountWithVat")); //(კორექტირების შემდეგ)
            formItems.Add("LinkTo", "amtACorE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "amtACorE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_amtACor");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("amtACorE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "amtTXACrS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("VatAmount")); //(კორექტირების შემდეგ)
            formItems.Add("LinkTo", "amtTXACrE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "amtTXACrE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_amtTXACr");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oForm.Items.Item("amtTXACrE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            //---------------------------------------------------

            left_s = 6;
            left_e = 127;
            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "taxInRcvdS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ReceiveVat"));
            formItems.Add("LinkTo", "taxInRcvdC");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            listValidValuesDict = new Dictionary<string, string>(); //ჩათვლილი (კომბო)
            listValidValuesDict.Add("", "");
            listValidValuesDict.Add("Y", BDOSResources.getTranslate("Yes"));
            listValidValuesDict.Add("N", BDOSResources.getTranslate("No"));

            formItems = new Dictionary<string, object>();
            itemName = "taxInRcvdC"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_TaxInRcvd");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Description", BDOSResources.getTranslate("VATReceived"));
            formItems.Add("ValidValues", listValidValuesDict);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "vatRDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("VATReceiveDate"));
            formItems.Add("LinkTo", "vatRDateE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "vatRDateE"; //10 characters 
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_vatRDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 2);
            formItems.Add("Width", width_e / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "wblS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s - 5);
            formItems.Add("Width", width_s * 2.5);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("LinkedAPInvoicesAndCreditNotes"));
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 12);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "addMTRB"; //10 characters 


            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_s);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Add"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "delMTRB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_s + 100 + 1);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Delete"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + top + 1;

            int wblMTRWidth = oForm.ClientWidth * 2 / 3;

            formItems = new Dictionary<string, object>();
            itemName = "wblMTR"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            formItems.Add("Left", left_s);
            formItems.Add("Width", wblMTRWidth);
            formItems.Add("Top", top);
            //formItems.Add("Height", 100);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("AffectsFormMode", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;

            multiSelection = false;
            objectType = "18";
            string uniqueID_lf_PurchaseInvoiceCFL = "PurchaseInvoice_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_PurchaseInvoiceCFL);
            objectType = "19";
            string uniqueID_lf_PurchaseInvoiceCreditMemoCFL = "PurchaseInvoiceCreditMemo_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_PurchaseInvoiceCreditMemoCFL);
            objectType = "204";
            string uniqueID_lf_DownPaymentInvoiceCFL = "DownPaymentInvoice_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_DownPaymentInvoiceCFL);

            oColumn = oColumns.Add("LineID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "  ";
            oColumn.Width = 20 - 1;
            oColumn.Editable = false;
            oColumn.Visible = true;

            wblMTRWidth = wblMTRWidth - 20 - 1;

            listValidValues = new List<string>();
            //listValidValues.Add("");
            listValidValues.Add(BDOSResources.getTranslate("APInvoice")); //0 //შესყიდვა
            listValidValues.Add(BDOSResources.getTranslate("APCreditNote")); //1 //შესყიდვის კორექტირება
            listValidValues.Add(BDOSResources.getTranslate("APDownPaymentInvoice")); //2 //გაცემული ავანსი

            oColumn = oColumns.Add("U_baseDocT", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocumentType");
            oColumn.Width = wblMTRWidth / 5;
            oColumn.Editable = true;
            oColumn.Visible = true;
            for (int i = 0; i < listValidValues.Count(); i++)
            {
                oColumn.ValidValues.Add(i == 0 & listValidValues[i] == "" ? "-1" : i.ToString(), listValidValues[i]);
            }

            oColumn = oColumns.Add("U_baseDoc", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Document");
            oColumn.Width = 1;
            oColumn.Editable = true;
            oColumn.Visible = false;

            oColumn = oColumns.Add("U_baseDTxt", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Document");
            oColumn.Width = wblMTRWidth / 5;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("U_amtBsDc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("AmountWithVat");
            oColumn.Width = wblMTRWidth / 5;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("U_tAmtBsDc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatAmount");
            oColumn.Width = wblMTRWidth / 5;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("U_wbNumber", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("WaybillNumber");
            oColumn.Width = wblMTRWidth / 5;
            oColumn.Editable = false;
            oColumn.Visible = true;

            SAPbouiCOM.DBDataSource oDBDataSource;
            oDBDataSource = oForm.DataSources.DBDataSources.Add("@BDO_TXR1");

            oColumn = oColumns.Item("LineID");
            oColumn.DataBind.SetBound(true, "@BDO_TXR1", "LineID");

            oColumn = oColumns.Item("U_baseDocT");
            oColumn.DataBind.SetBound(true, "@BDO_TXR1", "U_baseDocT");
            oColumn.DisplayDesc = true;
            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

            oColumn = oColumns.Item("U_baseDoc");
            oColumn.DataBind.SetBound(true, "@BDO_TXR1", "U_baseDoc");

            oColumn = oColumns.Item("U_baseDTxt");
            oColumn.DataBind.SetBound(true, "@BDO_TXR1", "U_baseDTxt");

            oColumn = oColumns.Item("U_amtBsDc");
            oColumn.DataBind.SetBound(true, "@BDO_TXR1", "U_amtBsDc");

            oColumn = oColumns.Item("U_tAmtBsDc");
            oColumn.DataBind.SetBound(true, "@BDO_TXR1", "U_tAmtBsDc");

            oColumn = oColumns.Item("U_wbNumber");
            oColumn.DataBind.SetBound(true, "@BDO_TXR1", "U_wbNumber");

            oMatrix.Clear();
            oDBDataSource.Query();
            oMatrix.LoadFromDataSource();

            left_s = 295;
            left_e = left_s + 121;

            int wblMTR2Width = oForm.ClientWidth / 4;

            formItems = new Dictionary<string, object>();
            itemName = "wblMTR2"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            formItems.Add("Left", oForm.ClientWidth - wblMTR2Width);
            formItems.Add("Width", wblMTR2Width);
            formItems.Add("Top", top);
            formItems.Add("Height", 100);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("AffectsFormMode", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR2").Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("LineID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "  ";
            oColumn.Width = 20;
            oColumn.Editable = false;
            oColumn.Visible = true;

            wblMTR2Width = wblMTR2Width - 20;

            oColumn = oColumns.Add("U_wbNumber", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("WaybillNumber");
            oColumn.Width = wblMTR2Width;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Item("LineID");
            oColumn.DataBind.SetBound(true, "@BDO_TXR2", "LineID");

            oColumn = oColumns.Item("U_wbNumber");
            oColumn.DataBind.SetBound(true, "@BDO_TXR2", "U_wbNumber");

            oMatrix.Clear();
            oDBDataSource.Query();
            oMatrix.LoadFromDataSource();

            left_s = 6;

            top = top + oForm.Items.Item("wblMTR").Height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "dpInvS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s - 5);
            formItems.Add("Width", width_s * 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("LinkedDPMTaxInvoices"));
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 12);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "addDPinv"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_s);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ChooseDownPaymentTaxInvoice"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "delDPinv"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_s + 100 + 1);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Delete"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + top + 1;

            int itemsWidth = oForm.ClientWidth * 2 / 3;

            formItems = new Dictionary<string, object>();
            itemName = "DPitems"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            formItems.Add("Left", left_s);
            formItems.Add("Width", itemsWidth);
            formItems.Add("Top", top);
            //formItems.Add("Height", 100);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("AffectsFormMode", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("DPitems").Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            oColumns = oMatrix.Columns;

            multiSelection = false;
            objectType = "5";
            string uniqueID_lf_VatType = "CFLvatType";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_VatType);

            //პირობის დადება
            oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_VatType);
            oCons = oCFL.GetConditions();
            oCon = oCons.Add();
            oCon.Alias = "Category";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "I";
            oCFL.SetConditions(oCons);

            oColumn = oColumns.Add("LineID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "  ";
            oColumn.Width = 20 - 1;
            oColumn.Editable = false;
            oColumn.Visible = true;

            itemsWidth = itemsWidth - 20 - 1;

            oColumn = oColumns.Add("goods", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Item");
            oColumn.Width = itemsWidth / 6;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("g_unit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomName");
            oColumn.Width = itemsWidth / 6;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("g_number", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Quantity");
            oColumn.Width = itemsWidth / 6;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("fullAmount", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("GrossAmount");
            oColumn.Width = itemsWidth / 6;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("drgAmount", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatAmount");
            oColumn.Width = itemsWidth / 6;
            oColumn.Editable = false;
            oColumn.Visible = true;

            //oColumn = oColumns.Add("vatType", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn = oColumns.Add("vatType", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VATCode");
            oColumn.Width = itemsWidth / 6;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oDBDataSource = oForm.DataSources.DBDataSources.Add("@BDO_TXR3");

            oColumn = oColumns.Item("goods");
            oColumn.DataBind.SetBound(true, "@BDO_TXR3", "U_goods");

            oColumn = oColumns.Item("g_unit");
            oColumn.DataBind.SetBound(true, "@BDO_TXR3", "U_g_unit");

            oColumn = oColumns.Item("g_number");
            oColumn.DataBind.SetBound(true, "@BDO_TXR3", "U_g_number");

            oColumn = oColumns.Item("fullAmount");
            oColumn.DataBind.SetBound(true, "@BDO_TXR3", "U_full_amount");

            oColumn = oColumns.Item("drgAmount");
            oColumn.DataBind.SetBound(true, "@BDO_TXR3", "U_drg_amount");

            oColumn = oColumns.Item("vatType");
            oColumn.DataBind.SetBound(true, "@BDO_TXR3", "U_vat_type");

            oColumn.ChooseFromListUID = uniqueID_lf_VatType;
            oColumn.ChooseFromListAlias = "Code";

            oMatrix.Clear();
            oDBDataSource.Query();
            oMatrix.LoadFromDataSource();

            formItems = new Dictionary<string, object>();
            itemName = "DPinvoices"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            formItems.Add("Left", left_s);
            formItems.Add("Width", itemsWidth);
            formItems.Add("Top", top);
            //formItems.Add("Height", 100);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("AffectsFormMode", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("DPinvoices").Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            oColumns = oMatrix.Columns;

            //multiSelection = false;
            //objectType = "5";
            //string uniqueID_lf_VatType = "Code";
            //FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_VatType);

            oColumn = oColumns.Add("LineID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "  ";
            oColumn.Width = 20 - 1;
            oColumn.Editable = false;
            oColumn.Visible = true;

            itemsWidth = itemsWidth - 20 - 1;

            //multiSelection = false;
            //objectType = "UDO_F_BDO_TAXR_D";
            //string uniqueID_lf_DPinvoiceCFL = "DPinvoiceCFL";
            //FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_DPinvoiceCFL);

            oColumn = oColumns.Add("DPinvoice", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("APDownPaymentInvoice");
            oColumn.Width = itemsWidth / 6;
            oColumn.Editable = true;
            oColumn.Visible = true;
            oColumn.Editable = false;

            oColumn = oColumns.Add("maxAmount", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("OpenVAT");
            oColumn.Width = itemsWidth / 6;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("drgAmount", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatAmount");
            oColumn.Width = itemsWidth / 6;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oDBDataSource = oForm.DataSources.DBDataSources.Add("@BDO_TXR4");

            oColumn = oColumns.Item("DPinvoice");
            oColumn.DataBind.SetBound(true, "@BDO_TXR4", "U_DP_invoice");

            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "UDO_F_BDO_TAXR_D";

            oColumn = oColumns.Item("maxAmount");
            oColumn.DataBind.SetBound(true, "@BDO_TXR4", "U_max_amount");

            oColumn = oColumns.Item("drgAmount");
            oColumn.DataBind.SetBound(true, "@BDO_TXR4", "U_drg_amount");

            oMatrix.Clear();
            oDBDataSource.Query();
            oMatrix.LoadFromDataSource();

            //სარდაფი
            left_s = 6;
            left_e = 127;
            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "CreatorS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Creator"));
            formItems.Add("LinkTo", "CreatorE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "CreatorE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "Creator");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item("CreatorE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "RemarksS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Remarks"));
            formItems.Add("LinkTo", "RemarksE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "RemarksE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "Remark");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e * 3);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ScrollBars", SAPbouiCOM.BoScrollBars.sb_Vertical);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "CommentS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Comment"));
            formItems.Add("LinkTo", "CommentE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "CommentE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXR");
            formItems.Add("Alias", "U_comment");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e * 3);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //ღილაკები

            left_s = 295;
            left_e = left_s + 121;

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "postB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_e - 100 - 5);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", 19);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Post"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("RSUpdateStatus"));
            listValidValuesDict.Add("deny", BDOSResources.getTranslate("RSDeny"));
            listValidValuesDict.Add("confirmation", BDOSResources.getTranslate("RSConfirm"));
            listValidValuesDict.Add("addToTheDeclaration", BDOSResources.getTranslate("RSAddDeclaration"));
            listValidValuesDict.Add("update", BDOSResources.getTranslate("RSUpdate"));

            formItems = new Dictionary<string, object>();
            itemName = "operationB";
            formItems.Add("Caption", BDOSResources.getTranslate("Operations"));
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
            formItems.Add("Left", left_e);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", 19);
            formItems.Add("UID", itemName);
            formItems.Add("ValidValues", listValidValuesDict);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("AffectsFormMode", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem = null;
            int height = 15;
            int top = 6;
            top = top + height + 1;

            oItem = oForm.Items.Item("cardCodeS");
            oItem.Top = top;
            oItem = oForm.Items.Item("cardCodeE");
            oItem.Top = top;
            oItem = oForm.Items.Item("cardCodeNE");
            oItem.Top = top;
            oItem = oForm.Items.Item("cardCodeLB");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("cardCodeTS");
            oItem.Top = top;
            oItem = oForm.Items.Item("cardCodeTE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("recvDateS");
            oItem.Top = top;
            oItem = oForm.Items.Item("recvDateE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("corrInvCH");
            oItem.Top = top;
            oItem = oForm.Items.Item("elctrnicCH");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("downPmntCH");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("BDOSJrnEnS");
            oItem.Top = top;
            oItem = oForm.Items.Item("BDOSJrnEnt");
            oItem.Top = top;
            oItem = oForm.Items.Item("BDOSJEntLB");
            oItem.Top = top;

            oItem = oForm.Items.Item("corrTypeCB");
            oItem.Top = top;

            top = 6;

            oItem = oForm.Items.Item("No.S");
            oItem.Top = top;
            oItem = oForm.Items.Item("SeriesC");
            oItem.Top = top;
            oItem = oForm.Items.Item("DocNumE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("StatusS");
            oItem.Top = top;
            oItem = oForm.Items.Item("StatusC");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("CreateDatS");
            oItem.Top = top;
            oItem = oForm.Items.Item("CreateDatE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("DocDateS");
            oItem.Top = top;
            oItem = oForm.Items.Item("DocDateE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("taxDateS");
            oItem.Top = top;
            oItem = oForm.Items.Item("taxDateE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("corrDocIDS");
            oItem.Top = top;
            oItem = oForm.Items.Item("corrDocIDE");
            oItem.Top = top;
            oItem = oForm.Items.Item("corrDocE");
            oItem.Top = top;
            oItem = oForm.Items.Item("corrDocLB");
            oItem.Top = top;

            //ელექტრონული ანგარიშ-ფაქტურა ---->

            top = 86 + 40;

            oItem = oForm.Items.Item("taxInvS");
            oItem.Top = top;
            top = top + 25;

            oItem = oForm.Items.Item("seriesS");
            oItem.Top = top;
            oItem = oForm.Items.Item("seriesE");
            oItem.Top = top;
            oItem = oForm.Items.Item("numberE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("statusS1");
            oItem.Top = top;
            oItem = oForm.Items.Item("statusCB");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("confInfoS");
            oItem.Top = top;
            oItem = oForm.Items.Item("confInfoE");
            oItem.Top = top;
            oItem = oForm.Items.Item("confInfoE1");
            oItem.Top = top;
            oItem = oForm.Items.Item("confInfoLB");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("confDateS");
            oItem.Top = top;
            oItem = oForm.Items.Item("confDateE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("declNumS");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("declDateS");
            oItem.Top = top;
            oItem = oForm.Items.Item("declDateE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("declNmberS");
            oItem.Top = top;
            oItem = oForm.Items.Item("declNmberE");
            oItem.Top = top;

            top = 86 + 40;
            top = top + 25;

            oItem = oForm.Items.Item("invIDS");
            oItem.Top = top;
            oItem = oForm.Items.Item("invIDE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("opDateS");
            oItem.Top = top;
            oItem = oForm.Items.Item("opDateE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("amountS");
            oItem.Top = top;
            oItem = oForm.Items.Item("amountE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("amountTXS");
            oItem.Top = top;
            oItem = oForm.Items.Item("amountTXE");
            oItem.Top = top;

            top = oForm.Items.Item("declNmberS").Top;

            oItem = oForm.Items.Item("declStatS");
            oItem.Top = top;
            oItem = oForm.Items.Item("declStatS1");
            oItem.Top = top;
            top = top + height + 1;

            //კორექტირებული ფაქტურის მონაცემები ---->

            top = 223 + 40;

            oItem = oForm.Items.Item("21_U_S");
            oItem.Top = top;
            top = top + 25;

            oItem = oForm.Items.Item("amtCorS");
            oItem.Top = top;
            oItem = oForm.Items.Item("amtCorE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("amtTXCorS");
            oItem.Top = top;
            oItem = oForm.Items.Item("amtTXCorE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("taxInRcvdS");
            oItem.Top = top;
            oItem = oForm.Items.Item("taxInRcvdC");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("vatRDateS");
            oItem.Top = top;
            oItem = oForm.Items.Item("vatRDateE");
            oItem.Top = top;
            top = top + height + 1;

            //მონაცემები კორექტირების შემდეგ ---->

            top = 223 + 40;

            oItem = oForm.Items.Item("24_U_S");
            oItem.Top = top;

            top = top + 25;

            oItem = oForm.Items.Item("amtACorS");
            oItem.Top = top;
            oItem = oForm.Items.Item("amtACorE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("amtTXACrS");
            oItem.Top = top;
            oItem = oForm.Items.Item("amtTXACrE");
            oItem.Top = top;
            top = top + 2 * height + 1;

            //oItem = oForm.Items.Item("vatRecvdCH");
            //oItem.Top = top;
            oItem = oForm.Items.Item("vatRDateS");
            oItem.Top = top;
            oItem = oForm.Items.Item("vatRDateE");
            oItem.Top = top;

            top = top + height + 5;

            oItem = oForm.Items.Item("wblS");
            oItem.Top = top;
            top = top + height + 5;

            oItem = oForm.Items.Item("addMTRB");
            oItem.Top = top;
            oItem = oForm.Items.Item("delMTRB");
            oItem.Top = top;
            top = top + height + 1;

            int wblMTRWidth = oForm.ClientWidth * 2 / 3;
            oItem = oForm.Items.Item("wblMTR");
            oItem.Top = top;
            oItem.Width = wblMTRWidth;
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;
            oColumn = oMatrix.Columns.Item("LineID");
            oColumn.Width = 20 - 1;
            wblMTRWidth = wblMTRWidth - 20 - 1;
            oColumn = oMatrix.Columns.Item("U_baseDocT");
            oColumn.Width = wblMTRWidth / 5;
            oColumn = oMatrix.Columns.Item("U_baseDoc");
            oColumn.Width = 1;
            oColumn = oMatrix.Columns.Item("U_baseDTxt");
            oColumn.Width = wblMTRWidth / 5;
            oColumn = oMatrix.Columns.Item("U_amtBsDc");
            oColumn.Width = wblMTRWidth / 5;
            oColumn = oMatrix.Columns.Item("U_tAmtBsDc");
            oColumn.Width = wblMTRWidth / 5;
            oColumn = oMatrix.Columns.Item("U_wbNumber");
            oColumn.Width = wblMTRWidth / 5;

            int wblMTR2Width = oForm.ClientWidth / 4;
            oItem = oForm.Items.Item("wblMTR2");
            oItem.Top = top;
            oItem.Width = wblMTR2Width;
            oItem.Left = oForm.ClientWidth - wblMTR2Width;
            oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR2").Specific));
            oColumns = oMatrix.Columns;
            oColumn = oMatrix.Columns.Item("LineID");
            oColumn.Width = 20;
            wblMTRWidth = wblMTRWidth - 20;
            oColumn = oMatrix.Columns.Item("U_wbNumber");
            oColumn.Width = wblMTRWidth;

            //DPitems

            top = top + oForm.Items.Item("wblMTR").Height + 5;
            int topForButtons = top;
            oItem = oForm.Items.Item("dpInvS");
            oItem.Top = top;
            top = top + height + 5;

            oItem = oForm.Items.Item("addDPinv");
            oItem.Top = top;
            oItem = oForm.Items.Item("delDPinv");
            oItem.Top = top;
            top = top + height + 1;

            int itemsWidth = oForm.ClientWidth * 2 / 3;
            oItem = oForm.Items.Item("DPitems");
            oItem.Top = top;
            oItem.Width = itemsWidth;
            oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("DPitems").Specific));
            oColumns = oMatrix.Columns;
            oColumn = oMatrix.Columns.Item("LineID");
            oColumn.Width = 20 - 1;
            itemsWidth = itemsWidth - 20 - 1;
            oColumn = oMatrix.Columns.Item("goods");
            oColumn.Width = itemsWidth / 6;

            oColumn = oMatrix.Columns.Item("g_unit");
            oColumn.Width = itemsWidth / 6;

            oColumn = oMatrix.Columns.Item("g_number");
            oColumn.Width = itemsWidth / 6;

            oColumn = oMatrix.Columns.Item("fullAmount");
            oColumn.Width = itemsWidth / 6;

            oColumn = oMatrix.Columns.Item("drgAmount");
            oColumn.Width = itemsWidth / 6;

            oColumn = oMatrix.Columns.Item("vatType");
            oColumn.Width = itemsWidth / 6;

            //DPinvoices
            int dpinvoicesWidth = oForm.ClientWidth * 2 / 3;
            oItem = oForm.Items.Item("DPinvoices");
            oItem.Top = top;
            oItem.Width = dpinvoicesWidth;
            oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("DPinvoices").Specific));
            oColumns = oMatrix.Columns;
            oColumn = oMatrix.Columns.Item("LineID");
            oColumn.Width = 20 - 1;
            itemsWidth = itemsWidth - 20 - 1;

            oColumn = oMatrix.Columns.Item("DPinvoice");
            oColumn.Width = itemsWidth / 3;

            oColumn = oMatrix.Columns.Item("maxAmount");
            oColumn.Width = itemsWidth / 3;

            oColumn = oMatrix.Columns.Item("drgAmount");
            oColumn.Width = itemsWidth / 3;

            //სარდაფი
            top = topForButtons + height * 2 + 5 + 1 + 100;

            oItem = oForm.Items.Item("CreatorS");
            oItem.Top = top;
            oItem = oForm.Items.Item("CreatorE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("RemarksS");
            oItem.Top = top;
            oItem = oForm.Items.Item("RemarksE");
            oItem.Top = top;
            top = top + height + 1;
            //top = top + 3 * height + 1;

            oItem = oForm.Items.Item("CommentS");
            oItem.Top = top;
            oItem = oForm.Items.Item("CommentE");
            oItem.Top = top;

            //ღილაკები
            int topTemp1 = oForm.Items.Item("CommentE").Top + height + 1;
            int topTemp2 = oForm.ClientHeight - 28;
            //ღილაკები
            top = topTemp2 > topTemp1 ? topTemp2 : topTemp1;

            oItem = oForm.Items.Item("1");
            oItem.Top = top;

            oItem = oForm.Items.Item("2");
            oItem.Top = top;

            oItem = oForm.Items.Item("operationB");
            oItem.Left = oForm.ClientWidth - 6 - oItem.Width;
            oItem.Top = top;
            oItem = oForm.Items.Item("postB");
            oItem.Left = oForm.ClientWidth - 6 - oItem.Width - 100 - 5;
            oItem.Top = top;
        }

        public static void setSizeForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            try
            {
                oForm.ClientHeight = oForm.Height * 5 + 40;//1400;
                oForm.ClientWidth = oForm.Width * 2; //1000;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void resizeForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                reArrangeFormItems(oForm);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Item oItem = null;
            oForm.Freeze(true);

            try
            {
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDO_TAXR");

                string docEntry = oDBDataSource.GetValue("DocEntry", 0).Trim();

                string invoiceStatus = oDBDataSource.GetValue("U_status", 0).Trim();

                bool docEntryIsEmpty = string.IsNullOrEmpty(docEntry);
                oForm.Items.Item("corrInvCH").Enabled = (docEntryIsEmpty == true);
                oForm.Items.Item("downPmntCH").Enabled = (docEntryIsEmpty == true);

                oForm.Items.Item("CreateDatE").Enabled = (docEntryIsEmpty == true);
                oForm.Items.Item("DocDateE").Enabled = (docEntryIsEmpty == true);
                //|| (invoiceStatus == "received" || invoiceStatus == "correctionReceived" || invoiceStatus == "cancellationProcess")
                //&& oDBDataSource.GetValue("U_JrnEnt", 0) == "");

                oForm.Items.Item("BDOSJrnEnt").Enabled = false;
                oForm.Items.Item("statusCB").Enabled = false;

                string jrnEntry = oDBDataSource.GetValue("U_JrnEnt", 0).Trim();
                bool docJrnEntryIsEmpty = string.IsNullOrEmpty(jrnEntry);

                string corrInv = oDBDataSource.GetValue("U_corrInv", 0).Trim();
                oItem = oForm.Items.Item("21_U_S");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("amtCorS");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("amtCorE");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("amtTXCorS");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("amtTXCorE");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("24_U_S");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("amtACorS");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("amtACorE");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("amtTXACrS");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("amtTXACrE");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("corrDocIDS");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("corrDocIDE");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("corrDocE");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("corrTypeCB");
                oItem.Visible = corrInv == "N" ? false : true;

                setValidValuesBtnCombo(oForm, out errorText);

                string elctrnic = oDBDataSource.GetValue("U_elctrnic", 0).Trim();
                if (elctrnic == "N")
                {
                    oItem = oForm.Items.Item("seriesE"); //სერია
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("numberE"); //ნომერი
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("invIDE"); //ID
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("confDateE"); //დადასტურების თარიღი
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("declNmberE"); //დეკლარაციის ნომერი
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("opDateE"); //ოპერაციის თვე
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("corrDocIDE"); //კორექტირების დოკუმენტის ID
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("corrDocE"); //კორექტირების დოკუმენტის DocEntry
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("corrTypeCB"); //კორექტირების მიზეზი
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("amountE"); //თანხა დღგ-ის ჩათვლით
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("amountTXE"); //დღგ-ის თანხა
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("amtCorE"); //თანხა დღგ-ის ჩათვლით (კორექტირებული)
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("amtTXCorE"); //დღგ-ის თანხა (კორექტირებული)
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("amtACorE"); //თანხა დღგ-ის ჩათვლით (კორექტირების შემდეგ)
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("amtTXACrE"); //დღგ-ის თანხა (კორექტირების შემდეგ)
                    oItem.Enabled = true;
                }
                else
                {
                    oItem = oForm.Items.Item("seriesE"); //სერია
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("numberE"); //ნომერი
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("invIDE"); //ID
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("confDateE"); //დადასტურების თარიღი
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("declNmberE"); //დეკლარაციის ნომერი
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("opDateE"); //ოპერაციის თვე
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("corrDocIDE"); //კორექტირების დოკუმენტის ID
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("corrDocE"); //კორექტირების დოკუმენტის DocEntry
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("corrTypeCB"); //კორექტირების მიზეზი
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("amountE"); //თანხა დღგ-ის ჩათვლით
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("amountTXE"); //დღგ-ის თანხა
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("amtCorE"); //თანხა დღგ-ის ჩათვლით (კორექტირებული)
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("amtTXCorE"); //დღგ-ის თანხა (კორექტირებული)
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("amtACorE"); //თანხა დღგ-ის ჩათვლით (კორექტირების შემდეგ)
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("amtTXACrE"); //დღგ-ის თანხა (კორექტირების შემდეგ)
                    oItem.Enabled = false;
                }

                oItem = oForm.Items.Item("operationB");

                if (invoiceStatus == "paper") // ქაღალდის               
                    oItem.Visible = false;
                else if (invoiceStatus != "")
                    oItem.Visible = true;

                string downPmnt = oDBDataSource.GetValue("U_downPaymnt", 0).Trim();
                bool enabledDPinvoices = (downPmnt != "Y" && corrInv != "Y" && (invoiceStatus == "confirmed" || invoiceStatus == "corrected"));

                oForm.Items.Item("DPitems").Visible = downPmnt == "Y";
                oForm.Items.Item("wblMTR2").Visible = downPmnt != "Y";
                oForm.Items.Item("DPinvoices").Visible = downPmnt != "Y";

                oForm.Items.Item("addDPinv").Visible = (docJrnEntryIsEmpty == true && enabledDPinvoices);
                oForm.Items.Item("delDPinv").Visible = (docJrnEntryIsEmpty == true && enabledDPinvoices);
                oForm.Items.Item("DPinvoices").Enabled = (docJrnEntryIsEmpty == true && enabledDPinvoices);
                oForm.Items.Item("DPitems").Enabled = (docJrnEntryIsEmpty == true);

                bool DPitemsEditable = (downPmnt == "Y" && (docJrnEntryIsEmpty == true) && elctrnic == "N");
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("DPitems").Specific));
                oMatrix.Columns.Item("goods").Editable = DPitemsEditable;
                oMatrix.Columns.Item("g_unit").Editable = DPitemsEditable;
                oMatrix.Columns.Item("g_number").Editable = DPitemsEditable;
                oMatrix.Columns.Item("fullAmount").Editable = DPitemsEditable;
                oMatrix.Columns.Item("drgAmount").Editable = DPitemsEditable;
                //SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("fullAmount");
                //oColumn.Editable = (docJrnEntryIsEmpty == true) && elctrnic == "N";

                oForm.Items.Item("dpInvS").Specific.Caption = downPmnt == "Y" ? BDOSResources.getTranslate("DownPaymentTaxInvoiceContent") : BDOSResources.getTranslate("LinkedDPMTaxInvoices");
                oForm.Items.Item("taxDateS").Visible = downPmnt == "Y";
                oForm.Items.Item("taxDateE").Visible = downPmnt == "Y";
                if (downPmnt == "Y" && docJrnEntryIsEmpty)
                {
                    oForm.Items.Item("taxDateE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                    oForm.Items.Item("postB").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                }
                else
                {
                    oForm.Items.Item("taxDateE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                    oForm.Items.Item("postB").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                oForm.Update();
                GC.Collect();
            }
        }

        public static void comboSelect(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {


                if (pVal.ItemUID == "taxInRcvdC")
                {
                    string vatRecvd = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_TaxInRcvd", 0).Trim();
                    if (vatRecvd == "Y")
                    {
                        SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("vatRDateE").Specific;
                        DateTime vatRDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                        oEditText.Value = vatRDate.ToString("yyyyMMdd");
                    }
                    else
                    {
                        SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("vatRDateE").Specific;
                        oEditText.Value = "";
                    }
                }


                SAPbouiCOM.ButtonCombo oButtonCombo = ((SAPbouiCOM.ButtonCombo)(oForm.Items.Item("operationB").Specific));

                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "operationB")
                    {
                        oForm.Freeze(true);
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "wblMTR" & pVal.ColUID == "U_baseDocT")
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));
                        SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
                        if (cellPos == null)
                        {
                            return;
                        }

                        if (pVal.ItemChanged == true)
                        {
                            oMatrix.Columns.Item("U_baseDTxt").Cells.Item(cellPos.rowIndex).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            SAPbouiCOM.EditText oEditText = oMatrix.Columns.Item("U_baseDTxt").Cells.Item(cellPos.rowIndex).Specific;

                            try
                            {
                                oEditText.Value = "";
                                oEditText = oMatrix.Columns.Item("U_baseDoc").Cells.Item(cellPos.rowIndex).Specific;
                                oEditText.Value = "0";
                            }
                            catch
                            { }
                            try
                            {
                                oEditText = oMatrix.Columns.Item("U_amtBsDc").Cells.Item(cellPos.rowIndex).Specific; //თანხა დღგ-ის ჩათვლით
                                oEditText.Value = "0";
                            }
                            catch
                            { }
                            try
                            {
                                oEditText = oMatrix.Columns.Item("U_tAmtBsDc").Cells.Item(cellPos.rowIndex).Specific; //დღგ-ის თანხა
                                oEditText.Value = "0";
                            }
                            catch
                            { }
                        }
                    }

                    if (pVal.ItemUID == "operationB")
                    {
                        string selectedOperation = null;
                        if (oButtonCombo.Selected != null)
                        {
                            selectedOperation = oButtonCombo.Selected.Value;
                        }
                        else
                        {
                            return;
                        }

                        oForm.Freeze(false);
                        oButtonCombo.Caption = BDOSResources.getTranslate("Operations");

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("ToCompleteOperationWriteDocument"));
                            return;
                        }
                        else if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationImpossibleInMode"));
                            return;
                        }

                        if (selectedOperation != null)
                        {
                            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
                            if (errorText != null)
                            {
                                return;
                            }

                            string su = rsSettings["SU"];
                            string sp = rsSettings["SP"];

                            TaxInvoice oTaxInvoice = new TaxInvoice(rsSettings["ProtocolType"]);

                            bool chek_service_user = oTaxInvoice.check_usr(su, sp, out errorText);
                            if (chek_service_user == false)
                            {
                                Program.uiApp.MessageBox(BDOSResources.getTranslate("ServiceUserPasswordNotCorrect"));
                                return;
                            }
                            string statusRS = null;

                            if (selectedOperation == "updateStatus") //სტატუსების განახლება
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("DocEntry", 0));
                                operationRS(oTaxInvoice, selectedOperation, docEntry, -1, new DateTime(), null, out statusRS, out errorText);
                                if (errorText != null)
                                {
                                    Program.uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("RSUpdateStatus") + " " + BDOSResources.getTranslate("DoneSuccessfully"));
                                }
                            }
                            else if (selectedOperation == "deny") //უარყოფა
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("DocEntry", 0));
                                operationRS(oTaxInvoice, selectedOperation, docEntry, -1, new DateTime(), null, out statusRS, out errorText);
                                if (errorText != null)
                                {
                                    Program.uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("RSDeny") + " " + BDOSResources.getTranslate("DoneSuccessfully"));
                                }
                            }
                            else if (selectedOperation == "confirmation") //დადასტურება
                            {

                                if (oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_docDate", 0) == "")
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("PostingDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                }
                                else
                                {
                                    int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("DocEntry", 0));
                                    operationRS(oTaxInvoice, selectedOperation, docEntry, -1, new DateTime(), null, out statusRS, out errorText);
                                    if (errorText != null)
                                    {
                                        Program.uiApp.MessageBox(errorText);
                                    }
                                    else
                                    {
                                        Program.uiApp.MessageBox(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("RSConfirm") + " " + BDOSResources.getTranslate("DoneSuccessfully"));
                                    }
                                }
                            }
                            else if (selectedOperation == "addToTheDeclaration") //დეკლარაციაში დამატება
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("DocEntry", 0));
                                operationRS(oTaxInvoice, selectedOperation, docEntry, -1, new DateTime(), null, out statusRS, out errorText);
                                if (errorText != null)
                                {
                                    Program.uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("RSAddDeclaration") + " " + BDOSResources.getTranslate("DoneSuccessfully"));
                                }
                            }
                            else if (selectedOperation == "update") //განახლება
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("DocEntry", 0));
                                operationRS(oTaxInvoice, selectedOperation, docEntry, -1, new DateTime(), null, out statusRS, out errorText);
                                if (errorText != null)
                                {
                                    Program.uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("RSUpdate") + " " + BDOSResources.getTranslate("DoneSuccessfully"));
                                }
                            }
                            else if (selectedOperation == "receive") //ჩათვლა
                            {
                                //int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("DocEntry", 0));
                                //DateTime DeclDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                                //receiveVAT( docEntry, DeclDate, out errorText);
                                //if (errorText != null)
                                //{
                                //    Program.uiApp.MessageBox(errorText);
                                //}
                                //else
                                //{
                                //    Program.uiApp.MessageBox("ოპერაცია \"ჩათვლა\" წარმატებით დასრულდა!");
                                //}
                            }
                            if (selectedOperation != null)
                            {
                                //if(!string.IsNullOrEmpty(journalEntryForTaxWithPostingDateMsg))
                                //{
                                //    Program.uiApp.MessageBox(journalEntryForTaxWithPostingDateMsg);
                                //    journalEntryForTaxWithPostingDateMsg = null;
                                //}
                                FormsB1.SimulateRefresh();
                            }
                        }
                    }
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

        public static void formDataLoad(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                string declNumber = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_declNumber", 0).Trim();
                SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("declStatS1").Specific;
                if (declNumber == "")
                {
                    oStaticText.Caption = BDOSResources.getTranslate("NotLinked");
                }
                else
                {
                    oStaticText.Caption = BDOSResources.getTranslate("Linked");
                }

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    string docEntry;
                    SAPbobsCOM.Documents oDocument;
                    SAPbouiCOM.ComboBox oComboBox = oMatrix.Columns.Item("U_baseDocT").Cells.Item(i).Specific;
                    docEntry = oMatrix.Columns.Item("U_baseDTxt").Cells.Item(i).Specific.value.ToString();

                    if (String.IsNullOrEmpty(docEntry) == false)
                    {
                        bool cancelled = false;
                        if (oComboBox.Value == "0") //შესყიდვა
                        {
                            oDocument = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                            oDocument.GetByKey(Convert.ToInt32(docEntry));
                            if (oDocument.Cancelled == SAPbobsCOM.BoYesNoEnum.tYES)
                                cancelled = true;
                        }
                        else if (oComboBox.Value == "1") //დაბრუნება
                        {
                            oDocument = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);
                            oDocument.GetByKey(Convert.ToInt32(docEntry));
                            if (oDocument.Cancelled == SAPbobsCOM.BoYesNoEnum.tYES)
                                cancelled = true;
                        }
                        if (cancelled == true)
                        {
                            oMatrix.CommonSetting.SetRowBackColor(i, FormsB1.getLongIntRGB(255, 48, 48));
                        }
                        else
                        {
                            oMatrix.CommonSetting.SetRowBackColor(i, -1);
                        }
                    }
                }

                string DocEntry = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("DocEntry", 0).Trim();

                setVisibleFormItems(oForm, out errorText);

                //// გატარებები
                //SAPbouiCOM.DBDataSource DocDBSourceTAXP = oForm.DataSources.DBDataSources.Item(0);
                //string Ref1 = DocDBSourceTAXP.GetValue("DocEntry", 0);
                //string Ref2 = "UDO_F_BDO_TAXR_D";

                //SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //string query = "SELECT " +
                //                "*  " +
                //                "FROM \"OJDT\"  " +
                //                "WHERE \"StornoToTr\" IS NULL " +
                //                "AND \"Ref1\" = '" + Ref1 + "' " +
                //                "AND \"Ref2\" = '" + Ref2 + "' ";
                //oRecordSet.DoQuery(query);

                //if (!oRecordSet.EoF)
                //{
                //    oForm.Items.Item("BDOSJrnEnt").Specific.Value = oRecordSet.Fields.Item("TransId").Value;
                //}
                //else
                //{
                //    oForm.Items.Item("BDOSJrnEnt").Specific.Value = "";
                //}
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

        public static void setValidValuesBtnCombo(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            try
            {
                Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
                string invoiceStatus = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_status", 0).Trim();
                SAPbouiCOM.Item oItem = oForm.Items.Item("operationB");

                if (invoiceStatus == "paper") // ქაღალდის
                {
                    oItem.Visible = false;
                }
                else if (invoiceStatus == "received" || invoiceStatus == "correctionReceived") // მიღებული (დასადასტურებელი) || კორექტირება მიღებული (დასადასტურებელი)
                {
                    oItem.Visible = true;
                    listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("RSUpdateStatus"));
                    listValidValuesDict.Add("deny", BDOSResources.getTranslate("RSDeny"));
                    listValidValuesDict.Add("confirmation", BDOSResources.getTranslate("RSConfirm"));
                    listValidValuesDict.Add("update", BDOSResources.getTranslate("RSUpdate"));
                }
                else if (invoiceStatus == "cancellationProcess") // გაუქმების პროცესში
                {
                    oItem.Visible = true;
                    listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("RSUpdateStatus"));
                    listValidValuesDict.Add("confirmation", BDOSResources.getTranslate("RSConfirm"));
                    listValidValuesDict.Add("update", BDOSResources.getTranslate("RSUpdate"));
                }
                else if (invoiceStatus == "confirmed" || invoiceStatus == "correctionConfirmed" || invoiceStatus == "corrected") // დადასტურებული || კორექტირება დადასტურებული || კორექტირებული
                {
                    oItem.Visible = true;
                    listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("RSUpdateStatus"));
                    if (oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_declNumber", 0).Trim() == "")
                    {
                        listValidValuesDict.Add("addToTheDeclaration", BDOSResources.getTranslate("RSAddDeclaration"));
                    }
                    listValidValuesDict.Add("update", BDOSResources.getTranslate("RSUpdate"));
                }
                else
                {
                    oItem.Visible = true;
                    listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("RSUpdateStatus"));
                    listValidValuesDict.Add("update", BDOSResources.getTranslate("RSUpdate"));
                }
                //if (oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_vatRecvd", 0).Trim() != "Y" & oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_LinkStatus", 0).Trim() == "1") // ფაქტურა არ არის ჩათვლილი & მიბმული
                //{
                //    listValidValuesDict.Add("receive", "ჩათვლა");
                //}

                if (oItem.Visible == true)
                {
                    SAPbouiCOM.ButtonCombo oButtonCombo = ((SAPbouiCOM.ButtonCombo)(oForm.Items.Item("operationB").Specific));
                    int count = oButtonCombo.ValidValues.Count;

                    for (int i = 0; i < count; i++)
                    {
                        oButtonCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    foreach (KeyValuePair<string, string> keyValue in listValidValuesDict)
                    {
                        oButtonCombo.ValidValues.Add(keyValue.Key, keyValue.Value);
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void matrixColumnSetCfl(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (pVal.ColUID == "U_baseDTxt")
                {
                    if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED & pVal.BeforeAction == true) || (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS & pVal.BeforeAction == false))
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));

                        SAPbouiCOM.ComboBox oComboBox = oMatrix.Columns.Item("U_baseDocT").Cells.Item(pVal.Row).Specific;
                        SAPbouiCOM.Column oColumn;

                        if (oComboBox.Value == "0") //შესყიდვა
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            oColumn.ChooseFromListUID = "PurchaseInvoice_CFL";
                            oColumn.ChooseFromListAlias = "DocEntry";
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "18"; //SAPbouiCOM.BoLinkedObject.lf_PurchaseInvoice
                        }
                        else if (oComboBox.Value == "1") //დაბრუნება
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            oColumn.ChooseFromListUID = "PurchaseInvoiceCreditMemo_CFL";
                            oColumn.ChooseFromListAlias = "DocEntry";
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "19"; //SAPbouiCOM.BoLinkedObject.lf_PurchaseInvoiceCreditMemo
                        }
                        else if (oComboBox.Value == "2") //გაცემული ავანსი
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            oColumn.ChooseFromListUID = "DownPaymentInvoice_CFL";
                            oColumn.ChooseFromListAlias = "DocEntry";
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "204";
                        }
                    }
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

        public static void itemPressed(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "elctrnicCH")
                    {
                        string elctrnic = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_elctrnic", 0).Trim();
                        if (elctrnic == "N")
                        {
                            oForm.DataSources.DBDataSources.Item("@BDO_TAXR").SetValue("U_status", 0, "paper");
                        }
                        else
                        {
                            oForm.DataSources.DBDataSources.Item("@BDO_TAXR").SetValue("U_status", 0, "empty");
                        }
                        setVisibleFormItems(oForm, out errorText);
                    }
                    else if (pVal.ItemUID == "vatRDateE")
                    {
                        SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("vatRDateE").Specific;
                        DateTime vatRDate = DateTime.ParseExact(oEditText.Value, "yyyyMMdd", null);
                        vatRDate = new DateTime(vatRDate.Year, vatRDate.Month, 1);
                        oEditText.Value = vatRDate.ToString("yyyyMMdd");
                    }
                    else if (pVal.ItemUID == "opDateE")
                    {
                        SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("opDateE").Specific;
                        DateTime opDate = DateTime.ParseExact(oEditText.Value, "yyyyMMdd", null);
                        opDate = new DateTime(opDate.Year, opDate.Month, 1);
                        oEditText.Value = opDate.ToString("yyyyMMdd");
                    }
                    else if (pVal.ItemUID == "declDateE")
                    {
                        SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("declDateE").Specific;
                        DateTime declDate = DateTime.ParseExact(oEditText.Value, "yyyyMMdd", null);
                        declDate = new DateTime(declDate.Year, declDate.Month, 1);
                        oEditText.Value = declDate.ToString("yyyyMMdd");
                    }

                    else if (pVal.ItemUID == "corrInvCH")
                    {
                        setVisibleFormItems(oForm, out errorText);
                    }
                    else if (pVal.ItemUID == "downPmntCH")
                    {
                        setVisibleFormItems(oForm, out errorText);

                        //SAPbouiCOM.Item oItem = oForm.Items.Item("wblMTR");
                        //oItem.Width = oForm.ClientWidth;

                        //int wblMTRWidth = oForm.ClientWidth;
                        //FormsB1.resetWidthMatrixColumns(oForm, "wblMTR", "LineID", wblMTRWidth);
                    }
                }
                else
                {
                    //if (pVal.ItemUID == "wblMTR" & pVal.ColUID == "U_baseDocT")
                    //{
                    //    string downPaymnt = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_downPaymnt", 0).Trim();
                    //    string corrInv = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_corrInv", 0).Trim();

                    //    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));
                    //    SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("U_baseDocT");

                    //    oForm.Freeze(true);
                    //    foreach (SAPbouiCOM.ValidValue oValidValue in oColumn.ValidValues)
                    //    {
                    //        if (oValidValue.Value == "0")
                    //            oMatrix.Columns.Item("U_baseDocT").ValidValues.Remove("0", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    //        if (oValidValue.Value == "1")
                    //            oMatrix.Columns.Item("U_baseDocT").ValidValues.Remove("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    //        if (oValidValue.Value == "2")
                    //            oMatrix.Columns.Item("U_baseDocT").ValidValues.Remove("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    //    }

                    //    if (downPaymnt == "Y")
                    //    {
                    //        oMatrix.Columns.Item("U_baseDocT").ValidValues.Add("2", BDOSResources.getTranslate("APDownPaymentInvoice"));
                    //    }
                    //    else if (corrInv == "Y")
                    //    {
                    //        oMatrix.Columns.Item("U_baseDocT").ValidValues.Add("1", BDOSResources.getTranslate("APCreditNote"));
                    //    }
                    //    else
                    //    {
                    //        oMatrix.Columns.Item("U_baseDocT").ValidValues.Add("0", BDOSResources.getTranslate("APInvoice"));
                    //        oMatrix.Columns.Item("U_baseDocT").ValidValues.Add("1", BDOSResources.getTranslate("APCreditNote"));
                    //    }
                    //    oForm.Freeze(false);
                    //}
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

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction, out string errorText, out bool BubbleEvent)
        {
            errorText = null;

            BubbleEvent = true;
            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (beforeAction == true)
                {
                    if (sCFL_ID == "PurchaseInvoice_CFL" || sCFL_ID == "PurchaseInvoiceCreditMemo_CFL" || sCFL_ID == "DownPaymentInvoice_CFL")
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));
                        SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
                        if (cellPos == null)
                        {
                            return;
                        }

                        oForm.Freeze(true);

                        SAPbouiCOM.EditText oEditText = oMatrix.Columns.Item("U_wbNumber").Cells.Item(cellPos.rowIndex).Specific;
                        string wbNumber = oEditText.Value;

                        if (sCFL_ID != null)
                        {
                            string cardCode = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_cardCode", 0).Trim();
                            int docEntry;

                            string docEntryStr = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("DocEntry", 0).Trim();
                            SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                            if (string.IsNullOrEmpty(docEntryStr))
                            {
                                SAPbouiCOM.Condition oCon = oCons.Add();
                                oCon.Alias = "CardCode";
                                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCon.CondVal = cardCode;
                            }
                            else
                            {
                                docEntry = Convert.ToInt32(docEntryStr);

                                SAPbobsCOM.CompanyService oCompanyService = null;
                                SAPbobsCOM.GeneralService oGeneralService = null;
                                SAPbobsCOM.GeneralData oGeneralData = null;
                                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                                oCompanyService = Program.oCompany.GetCompanyService();
                                oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXR_D");
                                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                                oGeneralParams.SetProperty("DocEntry", docEntry);
                                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                                string baseDocType;
                                switch (sCFL_ID)
                                {
                                    case "PurchaseInvoice_CFL": baseDocType = "0"; break;
                                    case "PurchaseInvoiceCreditMemo_CFL": baseDocType = "1"; break;
                                    case "DownPaymentInvoice_CFL": baseDocType = "2"; break;
                                    default: baseDocType = null; break;
                                }
                                DataTable baseDocs = getListBaseDoc(oGeneralData, wbNumber, baseDocType, docEntry, out errorText);

                                int docCount = baseDocs.Rows.Count;
                                if (docCount == 0)
                                {
                                    SAPbouiCOM.Condition oCon = oCons.Add();
                                    oCon.Alias = "DocEntry";
                                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                    oCon.CondVal = "";
                                }
                                else
                                {
                                    for (int i = 0; i < docCount; i++)
                                    {
                                        SAPbouiCOM.Condition oCon = oCons.Add();
                                        oCon.Alias = "DocEntry";
                                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                        oCon.CondVal = baseDocs.Rows[i]["DocEntry"].ToString();
                                        oCon.Relationship = (i == docCount - 1) ? SAPbouiCOM.BoConditionRelationship.cr_NONE : SAPbouiCOM.BoConditionRelationship.cr_OR;
                                    }
                                }
                            }
                            oCFL.SetConditions(oCons);
                        }
                        oForm.Freeze(false);
                    }

                    if (sCFL_ID == "DPinvoiceCFL")
                    {

                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("DPinvoices").Specific));
                        SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
                        if (cellPos == null)
                        {
                            return;
                        }

                        //oForm.Freeze(true);

                        //SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        //    SAPbouiCOM.Condition oCon = oCons.Add();
                        //    oCon.Alias = "CardCode";
                        //    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        //    oCon.CondVal = cardCode;

                        //oCFL.SetConditions(oCons);

                        //oColumn.ChooseFromListUID = "DPinvoiceCFL";
                        //oColumn.ChooseFromListAlias = "DocEntry";
                        //SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                        //oLink.LinkedObjectType = "UDO_F_BDO_TAXR_D"; 


                    }
                }
                else if (beforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "Users_CFL")
                        {
                            string userID = Convert.ToString(oDataTable.GetValue("USERID", 0));
                            string userName = Convert.ToString(oDataTable.GetValue("U_NAME", 0));

                            oForm.DataSources.DBDataSources.Item("@BDO_TAXR").SetValue("U_confInfo", 0, userID);
                            oForm.DataSources.DBDataSources.Item("@BDO_TAXR").SetValue("U_confInfN", 0, userName);

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                        else if (sCFL_ID == "BusinessPartner_CFL")
                        {
                            string businessPartnerCode = Convert.ToString(oDataTable.GetValue("CardCode", 0));
                            string businessPartnerName = Convert.ToString(oDataTable.GetValue("CardName", 0));
                            string businessPartnerTin = Convert.ToString(oDataTable.GetValue("LicTradNum", 0));

                            oForm.DataSources.DBDataSources.Item("@BDO_TAXR").SetValue("U_cardCode", 0, businessPartnerCode);
                            oForm.DataSources.DBDataSources.Item("@BDO_TAXR").SetValue("U_cardCodeN", 0, businessPartnerName);
                            oForm.DataSources.DBDataSources.Item("@BDO_TAXR").SetValue("U_cardCodeT", 0, businessPartnerTin);

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                        else if (sCFL_ID == "PurchaseInvoice_CFL" || sCFL_ID == "PurchaseInvoiceCreditMemo_CFL" || sCFL_ID == "DownPaymentInvoice_CFL")
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));
                            SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
                            if (cellPos == null)
                            {
                                return;
                            }

                            int docEntry = Convert.ToInt32(oDataTable.GetValue("DocEntry", 0));
                            double gTotal = 0;
                            double lineVat = 0;
                            string wbNumber = null;
                            SAPbobsCOM.Documents oBaseDoc = null;

                            if (sCFL_ID == "PurchaseInvoice_CFL")
                            {
                                APInvoice.getAmount(docEntry, out gTotal, out lineVat, out errorText);
                                oBaseDoc = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                if (oBaseDoc.GetByKey(docEntry) == true)
                                {
                                    wbNumber = oBaseDoc.UserFields.Fields.Item("U_BDO_WBNo").Value;
                                }
                            }
                            else if (sCFL_ID == "PurchaseInvoiceCreditMemo_CFL")
                            {
                                APCreditMemo.getAmount(docEntry, out gTotal, out lineVat, out errorText);
                                oBaseDoc = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);
                                if (oBaseDoc.GetByKey(docEntry) == true)
                                {
                                    wbNumber = oBaseDoc.UserFields.Fields.Item("U_BDO_WBNo").Value;
                                }
                            }
                            else if (sCFL_ID == "DownPaymentInvoice_CFL")
                            {
                                APDownPaymentInvoice.getAmount(docEntry, out gTotal, out lineVat, out errorText);
                            }

                            SAPbouiCOM.EditText oEditText;

                            try
                            {
                                oEditText = oMatrix.Columns.Item("U_baseDTxt").Cells.Item(cellPos.rowIndex).Specific;
                                oEditText.Value = docEntry.ToString();
                            }
                            catch
                            { }
                            try
                            {
                                oEditText = oMatrix.Columns.Item("U_baseDoc").Cells.Item(cellPos.rowIndex).Specific;
                                oEditText.Value = docEntry.ToString();
                            }
                            catch
                            { }
                            try
                            {
                                oEditText = oMatrix.Columns.Item("U_amtBsDc").Cells.Item(cellPos.rowIndex).Specific; //თანხა დღგ-ის ჩათვლით
                                oEditText.Value = FormsB1.ConvertDecimalToString(Convert.ToDecimal(gTotal));
                            }
                            catch
                            { }
                            try
                            {
                                oEditText = oMatrix.Columns.Item("U_tAmtBsDc").Cells.Item(cellPos.rowIndex).Specific; //დღგ-ის თანხა
                                oEditText.Value = FormsB1.ConvertDecimalToString(Convert.ToDecimal(lineVat));
                            }
                            catch
                            { }
                            try
                            {
                                oEditText = oMatrix.Columns.Item("U_wbNumber").Cells.Item(cellPos.rowIndex).Specific;
                                oEditText.Value = wbNumber;
                            }
                            catch
                            { }

                        }
                        else if (sCFL_ID == "DPinvoiceCFL")
                        {
                            BubbleEvent = false;

                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("DPinvoices").Specific));
                            SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
                            if (cellPos == null)
                            {
                                return;
                            }

                            int docEntry = Convert.ToInt32(oDataTable.GetValue("DocEntry", 0));
                            decimal gTotal = 0;

                            gTotal = Convert.ToDecimal(oDataTable.GetValue("U_amountTX", 0));

                            SAPbouiCOM.EditText oEditText;

                            try
                            {
                                oEditText = oMatrix.Columns.Item("DPinvoice").Cells.Item(cellPos.rowIndex).Specific;
                                oEditText.Value = docEntry.ToString();
                            }
                            catch
                            { }
                            try
                            {
                                oEditText = oMatrix.Columns.Item("maxAmount").Cells.Item(cellPos.rowIndex).Specific;
                                oEditText.Value = FormsB1.ConvertDecimalToString(gTotal);
                            }
                            catch
                            { }
                            try
                            {
                                oEditText = oMatrix.Columns.Item("drgAmount").Cells.Item(cellPos.rowIndex).Specific;
                                oEditText.Value = FormsB1.ConvertDecimalToString(gTotal);
                            }
                            catch
                            { }

                        }
                        else if (sCFL_ID == "CFLvatType")
                        {

                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("DPitems").Specific));
                            SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
                            if (cellPos == null)
                            {
                                return;
                            }

                            string Code = oDataTable.GetValue("Code", 0);

                            SAPbouiCOM.EditText oEditText;

                            try
                            {
                                oEditText = oMatrix.Columns.Item("vatType").Cells.Item(cellPos.rowIndex).Specific;
                                oEditText.Value = Code;
                            }
                            catch
                            { }

                        }
                    }

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                    setVisibleFormItems(oForm, out errorText);
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

        public static string getStatusValueByStatusNumber(string statusNumber)
        {
            if (statusNumber == "1")
                return "received"; //მიღებული (დასადასტურებელი)
            else if (statusNumber == "2")
                return "confirmed"; //დადასტურებული
            else if (statusNumber == "3")
                return "corrected"; //კორექტირებული
            else if (statusNumber == "4")
                return "correctionDenied"; //უარყოფილი კორექტირებული
            else if (statusNumber == "5")
                return "correctionReceived"; //კორექტირება მიღებული (დასადასტურებელი)
            else if (statusNumber == "6")
                return "cancellationProcess"; //გაუქმების პროცესში
            else if (statusNumber == "7")
                return "canceled"; //გაუქმებული
            else if (statusNumber == "8")
                return "correctionConfirmed"; //კორექტირება დადასტურებული
            else if (statusNumber == "9")
                return "replaced"; //ჩანაცვლებული
            else if (statusNumber == "0") //უარყოფილი;
                return "denied";
            else if (statusNumber == "-1") //წაშლილი;
                return "removed";
            else
                return "empty";
        }

        /// <summary></summary>
        /// <param name="Program.oCompany"></param>
        /// <param name="docEntry"></param>
        /// <param name="baseDocType"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public static Dictionary<string, object> getTaxInvoiceReceivedDocumentInfo(int docEntry, string baseDocType, string cardCode, out string errorText)
        {
            errorText = null;

            Dictionary<string, object> taxDocInfo = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT 
            ""BDO_TAXR"".""CreateDate"" AS ""createDate"",
            ""BDO_TAXR"".""U_opDate"" AS ""opDate"",
            ""BDO_TAXR"".""DocEntry"" AS ""docEntry"",
            ""BDO_TAXR"".""U_invID"" AS ""invID"",
            ""BDO_TAXR"".""U_number"" AS ""number"",
            ""BDO_TAXR"".""U_series"" AS ""series"",
            ""BDO_TAXR"".""U_status"" AS ""status"",
            ""BDO_TAXR"".""U_cardCodeT"" AS ""cardCodeT""
            FROM ""@BDO_TAXR"" AS ""BDO_TAXR"" 
            INNER JOIN ""@BDO_TXR1"" AS ""BDO_TXR1"" 
            ON ""BDO_TXR1"".""DocEntry"" = ""BDO_TAXR"".""DocEntry"" 
            WHERE ""BDO_TXR1"".""U_baseDoc"" = '" + docEntry + @"' AND ""BDO_TXR1"".""U_baseDocT"" = '" + baseDocType +
            @"' AND ""BDO_TAXR"".""U_cardCode"" = N'" + cardCode +
            @"' AND ""BDO_TAXR"".""Canceled"" = 'N' AND (""BDO_TAXR"".""U_status"" NOT IN ('removed', 'canceled') OR ""BDO_TAXR"".""U_status"" IS NULL)";

            try
            {
                oRecordSet.DoQuery(query);
                if (oRecordSet.RecordCount > 0)
                {
                    taxDocInfo = new Dictionary<string, object>();

                    while (!oRecordSet.EoF)
                    {
                        taxDocInfo.Add("docEntry", oRecordSet.Fields.Item("docEntry").Value);
                        taxDocInfo.Add("invID", oRecordSet.Fields.Item("invID").Value);
                        taxDocInfo.Add("number", oRecordSet.Fields.Item("number").Value);
                        taxDocInfo.Add("series", oRecordSet.Fields.Item("series").Value);
                        taxDocInfo.Add("status", oRecordSet.Fields.Item("status").Value);
                        taxDocInfo.Add("createDate", oRecordSet.Fields.Item("createDate").Value.ToString("yyyyMMdd"));
                        taxDocInfo.Add("opDate", oRecordSet.Fields.Item("opDate").Value.ToString("yyyyMMdd"));

                        oRecordSet.MoveNext();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return taxDocInfo;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }

            return taxDocInfo;
        }

        public static Dictionary<string, object> getTaxInvoiceReceivedDocumentInfo(int docEntry, out string errorText)
        {
            errorText = null;

            Dictionary<string, object> taxDocInfo = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT 
            ""BDO_TAXR"".""CreateDate"" AS ""createDate"",
            ""BDO_TAXR"".""U_opDate"" AS ""opDate"",
            ""BDO_TAXR"".""DocEntry"" AS ""docEntry"",
            ""BDO_TAXR"".""U_invID"" AS ""invID"",
            ""BDO_TAXR"".""U_number"" AS ""number"",
            ""BDO_TAXR"".""U_series"" AS ""series"",
            ""BDO_TAXR"".""U_status"" AS ""status"",
            ""BDO_TAXR"".""U_cardCodeT"" AS ""cardCodeT""
            FROM ""@BDO_TAXR"" AS ""BDO_TAXR"" 
           
            WHERE ""BDO_TAXR"".""DocEntry"" = '" + docEntry + @"'";

            try
            {
                oRecordSet.DoQuery(query);
                if (oRecordSet.RecordCount > 0)
                {
                    taxDocInfo = new Dictionary<string, object>();

                    while (!oRecordSet.EoF)
                    {
                        taxDocInfo.Add("docEntry", oRecordSet.Fields.Item("docEntry").Value);
                        taxDocInfo.Add("invID", oRecordSet.Fields.Item("invID").Value);
                        taxDocInfo.Add("number", oRecordSet.Fields.Item("number").Value);
                        taxDocInfo.Add("series", oRecordSet.Fields.Item("series").Value);
                        taxDocInfo.Add("status", oRecordSet.Fields.Item("status").Value);
                        taxDocInfo.Add("createDate", oRecordSet.Fields.Item("createDate").Value.ToString("yyyyMMdd"));
                        taxDocInfo.Add("opDate", oRecordSet.Fields.Item("opDate").Value.ToString("yyyyMMdd"));

                        oRecordSet.MoveNext();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return taxDocInfo;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }

            return taxDocInfo;
        }

        public static void getPrimaryBaseDoc(int docEntry, string baseDocType, string cardCode, out List<int> baseDocList, out string errorText)
        {
            errorText = null;
            int corrDoc = 0;
            string corrInv = null;
            baseDocList = new List<int>();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT 
            ""BDO_TAXR"".""CreateDate"" AS ""createDate"",
            ""BDO_TAXR"".""DocEntry"" AS ""docEntry"",
            ""BDO_TAXR"".""U_invID"" AS ""invID"",
            ""BDO_TAXR"".""U_number"" AS ""number"",
            ""BDO_TAXR"".""U_series"" AS ""series"",
            ""BDO_TAXR"".""U_status"" AS ""status"",
            ""BDO_TAXR"".""U_cardCodeT"" AS ""cardCodeT"",
            ""BDO_TAXR"".""U_corrInv"" AS ""corrInv"",            
            ""BDO_TAXR"".""U_corrDoc"" AS ""corrDoc"",             
            ""BDO_TXR1"".""U_baseDoc"" AS ""baseDoc"",
            ""BDO_TXR1"".""U_baseDocT"" AS ""baseDocT""
            FROM ""@BDO_TAXR"" AS ""BDO_TAXR"" 
            INNER JOIN ""@BDO_TXR1"" AS ""BDO_TXR1"" 
            ON ""BDO_TXR1"".""DocEntry"" = ""BDO_TAXR"".""DocEntry"" 
            WHERE ""BDO_TAXR"".""U_cardCode"" = N'" + cardCode +
            @"' AND ""BDO_TAXR"".""DocEntry"" = '" + docEntry + "'";
            //@"' AND (""BDO_TAXR"".""Canceled"" = 'N' AND ""BDO_TAXR"".""U_status"" NOT IN ('removed', 'canceled'))";

            if (baseDocType == "2")
                query = query + @" AND ""BDO_TXR1"".""U_baseDocT"" = '" + baseDocType + "'";
            else
            {
                query = query + @" AND ((""BDO_TXR1"".""U_baseDocT"" = '0' AND ""BDO_TAXR"".""U_corrInv"" = 'N')";
                query = query + @" OR (""BDO_TXR1"".""U_baseDocT"" = '1' AND ""BDO_TAXR"".""U_corrInv"" = 'Y'))";
            }

            try
            {
                oRecordSet.DoQuery(query);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        corrInv = oRecordSet.Fields.Item("corrInv").Value.ToString();
                        if (corrInv == "Y")
                            corrDoc = Convert.ToInt32(oRecordSet.Fields.Item("corrDoc").Value);
                        else
                            baseDocList.Add(Convert.ToInt32(oRecordSet.Fields.Item("baseDoc").Value));

                        oRecordSet.MoveNext();
                    }
                    if (corrInv == "Y" && corrDoc != 0)
                    {
                        getPrimaryBaseDoc(corrDoc, baseDocType, cardCode, out baseDocList, out errorText);
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
        }

        private static DataTable getListBaseDoc(SAPbobsCOM.GeneralData oGeneralData, string overhead_no, string baseDocType, int docEntryTaxInv, out string errorText)
        {
            errorText = null;

            DataTable baseDocs = new DataTable();
            baseDocs.Columns.Add("DocEntry", typeof(int));
            baseDocs.Columns.Add("DocDate", typeof(DateTime));
            baseDocs.Columns.Add("BaseDocType", typeof(string));
            baseDocs.Columns.Add("GTotal", typeof(decimal));
            baseDocs.Columns.Add("LineVat", typeof(decimal));

            bool corrInv = oGeneralData.GetProperty("U_corrInv") == "Y" ? true : false;
            bool downPaymnt = oGeneralData.GetProperty("U_downPaymnt") == "Y" ? true : false;
            int corrDoc = oGeneralData.GetProperty("U_corrDoc");
            string cardCode = oGeneralData.GetProperty("U_cardCode");
            DateTime opDate = oGeneralData.GetProperty("U_opDate");
            DateTime firstDayMonth = new DateTime(opDate.Year, opDate.Month, 1);
            DateTime lastDayMonth = firstDayMonth.AddMonths(1).AddDays(-1);

            decimal amount = Convert.ToDecimal(oGeneralData.GetProperty("U_amount"));
            decimal amountTX = Convert.ToDecimal(oGeneralData.GetProperty("U_amountTX"));
            if (corrInv)
            {
                amount = Convert.ToDecimal(oGeneralData.GetProperty("U_amtACor"));
                amountTX = Convert.ToDecimal(oGeneralData.GetProperty("U_amtTXACr"));
            }

            string baseDocTable;
            string baseDocRowTable;
            bool fromDoc = true;

            if (amount < 0)
            {
                amount = amount * (-1);
            }
            if (amountTX < 0)
            {
                amountTX = amountTX * (-1);
            }

            if (string.IsNullOrEmpty(baseDocType))
            {
                fromDoc = false;
                if (downPaymnt)
                    baseDocType = "2"; //A/P DownPaymentInvoice
                else if (corrInv)
                    baseDocType = "1"; //A/P CreditNote
                else
                    baseDocType = "0"; //A/P Invoice
            }

            if (baseDocType == "0") //A/P Invoice
            {
                baseDocTable = "OPCH";
                baseDocRowTable = "PCH1";
            }
            else if (baseDocType == "1") //A/P CreditNote
            {
                baseDocTable = "ORPC";
                baseDocRowTable = "RPC1";
            }
            else if (baseDocType == "2") //A/P DownPaymentInvoice
            {
                baseDocTable = "ODPO";
                baseDocRowTable = "DPO1";
            }
            else
                return baseDocs;

            List<int> primaryBaseDocList = new List<int>();
            List<int> connectedDocList = new List<int>();
            if (corrInv)
            {
                if (downPaymnt)
                {
                    getPrimaryBaseDoc(corrDoc, baseDocType, cardCode, out primaryBaseDocList, out errorText);
                    connectedDocList = APDownPaymentInvoice.getAllConnectedDoc(primaryBaseDocList, out errorText);
                }
                else
                {
                    getPrimaryBaseDoc(corrDoc, null, cardCode, out primaryBaseDocList, out errorText);
                    connectedDocList = APCreditMemo.getAllConnectedDoc(primaryBaseDocList, "18", out errorText);
                }
            }

            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT
	             ""TABL"".""DocEntry"",
	             ""TABL"".""DocDate"",
	             '" + baseDocType + @"' AS ""BaseDocType"",
	             SUM(""TBL1"".""GTotal"") AS ""GTotal"",
            	 SUM(""TBL1"".""LineVat"") AS ""LineVat"" 
            FROM """ + baseDocTable + @""" AS ""TABL"" 
            LEFT JOIN """ + baseDocRowTable + @""" AS ""TBL1"" ON ""TBL1"".""DocEntry"" = ""TABL"".""DocEntry"""
            +
            (baseDocType == "0" ? @"LEFT JOIN ""OPDN"" ON ""TBL1"".""BaseEntry"" = ""OPDN"".""DocEntry""" : "")
            +

            @"WHERE ""TABL"".""CANCELED"" = 'N' 
            AND ""TABL"".""CardCode"" = N'" + cardCode + @"'
            AND ""TABL"".""DocDate"" >= '" + firstDayMonth.ToString("yyyyMMdd") + @"' AND ""TABL"".""DocDate"" <= '" + lastDayMonth.ToString("yyyyMMdd") + "'";

            if (baseDocType != "2" && string.IsNullOrEmpty(overhead_no) == false) //ვეძებთ ზედნადების ნომრით
            {
                query = query + @" AND (""TABL"".""U_BDO_WBNo"" = '" + overhead_no + "'" +
                    (baseDocType == "0" ? @" OR ""OPDN"".""U_BDO_WBNo""= '" + overhead_no + "' " : "") + ") ";
            }
            if (baseDocType == "2")
            {
                query = query + @" AND ""TABL"".""Posted"" = 'Y'";
            }
            if (corrInv && connectedDocList.Count > 0)
            {
                query = query + @" AND ""TABL"".""DocEntry"" IN (" + string.Join(",", connectedDocList) + ")";
            }
            query = query + @" AND ""TABL"".""DocEntry"" NOT IN ( SELECT
            	 ""BDO_TXR1"".""U_baseDoc"" 
            	FROM ""@BDO_TAXR"" AS ""BDO_TAXR"" 
            	INNER JOIN ""@BDO_TXR1"" AS ""BDO_TXR1"" ON ""BDO_TXR1"".""DocEntry"" = ""BDO_TAXR"".""DocEntry"" 
            	WHERE ""BDO_TAXR"".""Canceled"" = 'N' 
            		AND (""BDO_TAXR"".""U_status"" NOT IN ('removed',
            	 'canceled') OR ""BDO_TAXR"".""U_status"" IS NULL) 
            	AND ""BDO_TXR1"".""U_baseDocT"" = '" + baseDocType + @"' 
            	AND ""BDO_TAXR"".""U_cardCode"" = N'" + cardCode + @"'  
            	AND ""BDO_TXR1"".""DocEntry"" <> '" + docEntryTaxInv + @"' ) 
            GROUP BY ""TABL"".""DocEntry"",
            	 ""TABL"".""DocDate""";

            if (string.IsNullOrEmpty(overhead_no) && fromDoc == false) //ვეძებთ თანხით
            {
                decimal allowableDeviation = Convert.ToDecimal(CommonFunctions.getOADM("U_BDOSAllDev").ToString());
                decimal amountMax = amount + allowableDeviation;
                decimal amountMin = amount - allowableDeviation;
                query = query + @" HAVING(SUM(""TBL1"".""GTotal"") >= '" + amountMin.ToString(Nfi) + @"' AND SUM(""TBL1"".""GTotal"") <= '" + amountMax.ToString(Nfi) + "')";
            }

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    DataRow dataRow = baseDocs.Rows.Add();
                    dataRow["DocEntry"] = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value);
                    dataRow["DocDate"] = Convert.ToDateTime(oRecordSet.Fields.Item("DocDate").Value);
                    dataRow["BaseDocType"] = Convert.ToString(oRecordSet.Fields.Item("BaseDocType").Value);
                    dataRow["GTotal"] = Convert.ToDecimal(oRecordSet.Fields.Item("GTotal").Value);
                    dataRow["LineVat"] = Convert.ToDecimal(oRecordSet.Fields.Item("LineVat").Value);

                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return baseDocs;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }

            return baseDocs;
        }

        public static void addBaseDoc(int docEntry, int baseDocEntry, string baseDocType, string wbNumber, double baseDocGTotal, double baseDocLineVat, out string errorText)
        {
            errorText = null;
            try
            {
                SAPbobsCOM.CompanyService oCompanyService = null;
                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.GeneralData oGeneralData = null;
                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                oCompanyService = Program.oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXR_D");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                //Get UDO record
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", docEntry);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                //Create data for a row in the child table
                SAPbobsCOM.GeneralDataCollection oChildren = null;

                oChildren = oGeneralData.Child("BDO_TXR1");

                if (String.IsNullOrEmpty(wbNumber) == true)
                {
                    SAPbobsCOM.GeneralData oChild = null;
                    oChild = oChildren.Add();
                    oChild.SetProperty("U_baseDoc", baseDocEntry);
                    oChild.SetProperty("U_baseDTxt", baseDocEntry.ToString());
                    oChild.SetProperty("U_baseDocT", baseDocType);
                    oChild.SetProperty("U_amtBsDc", baseDocGTotal); //საფუძველი დოკუმენტის თანხა
                    oChild.SetProperty("U_tAmtBsDc", baseDocLineVat); //საფუძველი დოკუმენტის დღგ-ის თანხა
                }
                else
                {
                    bool find = false;
                    foreach (SAPbobsCOM.GeneralData oChild in oChildren)
                    {
                        if (wbNumber == oChild.GetProperty("U_wbNumber"))
                        {
                            oChild.SetProperty("U_baseDoc", baseDocEntry);
                            oChild.SetProperty("U_baseDTxt", baseDocEntry.ToString());
                            oChild.SetProperty("U_baseDocT", baseDocType);
                            oChild.SetProperty("U_amtBsDc", baseDocGTotal); //საფუძველი დოკუმენტის თანხა
                            oChild.SetProperty("U_tAmtBsDc", baseDocLineVat); //საფუძველი დოკუმენტის დღგ-ის თანხა
                            find = true;
                            break;
                        }
                    }
                    if (find == false)
                    {
                        SAPbobsCOM.GeneralData oChild = null;
                        oChild = oChildren.Add();
                        oChild.SetProperty("U_baseDoc", baseDocEntry);
                        oChild.SetProperty("U_baseDTxt", baseDocEntry.ToString());
                        oChild.SetProperty("U_baseDocT", baseDocType);
                        oChild.SetProperty("U_amtBsDc", baseDocGTotal); //საფუძველი დოკუმენტის თანხა
                        oChild.SetProperty("U_tAmtBsDc", baseDocLineVat); //საფუძველი დოკუმენტის დღგ-ის თანხა
                    }
                }

                oGeneralService.Update(oGeneralData);
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("FailedToAttachInvoiceDocument") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static List<string> getListTaxInvoiceReceived(string cardCode, string wbNumber, string baseDocType, DateTime docDate, out string errorText)
        {
            errorText = null;

            List<string> taxInvoiceDocList = new List<string>();

            DateTime opDate = docDate;
            DateTime firstDayMonth = new DateTime(opDate.Year, opDate.Month, 1);
            DateTime lastDayMonth = firstDayMonth.AddMonths(1).AddDays(-1);

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT 
            ""BDO_TAXR"".""DocEntry"" AS ""docEntry""
            FROM ""@BDO_TAXR"" AS ""BDO_TAXR""
            LEFT JOIN ""@BDO_TXR1"" AS ""BDO_TXR1""
            ON ""BDO_TXR1"".""DocEntry"" = ""BDO_TAXR"".""DocEntry"" 
            WHERE ""BDO_TAXR"".""U_cardCode"" = N'" + cardCode + "' " +
                       @"AND ""BDO_TAXR"".""U_status"" NOT IN ('removed', 'canceled', 'denied') AND ""BDO_TAXR"".""Canceled"" = 'N' 
                       AND ""BDO_TAXR"".""U_opDate"" >= '" + firstDayMonth.ToString("yyyyMMdd") + @"' AND ""BDO_TAXR"".""U_opDate"" <= '" + lastDayMonth.ToString("yyyyMMdd") + "' ";

            if (string.IsNullOrEmpty(wbNumber) == false)
            {
                query = query + @" AND ""BDO_TXR1"".""U_wbNumber"" = '" + wbNumber + "' ";
            }
            if (baseDocType == "0") //A/P Invoice
            {
                query = query + @" AND ""BDO_TAXR"".""U_corrInv"" = 'N' AND ""BDO_TAXR"".""U_downPaymnt"" = 'N' ";
            }
            else if (baseDocType == "1") //A/P CreditNote
            {
                //query = query + @" AND ""BDO_TAXR"".""U_corrInv"" = 'Y' AND ""BDO_TAXR"".""U_downPaymnt"" = 'N' ";
                query = query + @" AND ""BDO_TAXR"".""U_downPaymnt"" = 'N' ";
            }
            else if (baseDocType == "2") //A/P DownPaymentInvoice
            {
                query = query + @" AND ""BDO_TAXR"".""U_downPaymnt"" = 'Y' ";
            }

            query = query + @" GROUP BY ""BDO_TAXR"".""DocEntry"" ";

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    taxInvoiceDocList.Add(oRecordSet.Fields.Item("docEntry").Value.ToString());

                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return taxInvoiceDocList;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }

            return taxInvoiceDocList;
        }

        public static void removeBaseDoc(int docEntry, int baseDocEntry, string baseDocType, out string errorText)
        {
            errorText = null;
            try
            {
                SAPbobsCOM.CompanyService oCompanyService = null;
                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.GeneralData oGeneralData = null;
                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                oCompanyService = Program.oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXR_D");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                //Get UDO record
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", docEntry);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                //Create data for a row in the child table
                SAPbobsCOM.GeneralDataCollection oChildren = null;

                oChildren = oGeneralData.Child("BDO_TXR1");
                int i = 0;
                foreach (SAPbobsCOM.GeneralData oChild in oChildren)
                {
                    if (baseDocEntry == oChild.GetProperty("U_baseDoc") & baseDocType == oChild.GetProperty("U_baseDocT"))
                    {
                        oChildren.Remove(i);
                        break;
                    }
                    i = i + 1;
                }

                oGeneralService.Update(oGeneralData);
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
                GC.Collect();
            }
        }

        public static void addMatrixRow(SAPbouiCOM.Form oForm, string ItemUID, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            try
            {

                if (ItemUID == "addMTRB")
                {

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));
                    int index = 0;
                    if (oMatrix.RowCount == 0)
                    {
                        index = 1;
                    }
                    else
                    {
                        index = Convert.ToInt32(oMatrix.Columns.Item("LineID").Cells.Item(oMatrix.RowCount).Specific.Value) + 1;
                    }

                    oMatrix.AddRow(1, -1);
                    oMatrix.AutoResizeColumns();
                    oMatrix.Columns.Item("LineID").Cells.Item(oMatrix.RowCount).Specific.Value = index;
                    SAPbouiCOM.ComboBox oComboBox = oMatrix.Columns.Item("U_baseDocT").Cells.Item(oMatrix.RowCount).Specific;
                    string downPaymnt = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_downPaymnt", 0).Trim();
                    string corrInv = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_corrInv", 0).Trim();
                    if (downPaymnt == "Y")
                        //oMatrix.Columns.Item("U_baseDocT").Cells.Item(oMatrix.RowCount).Specific.Value = "2";
                        oComboBox.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    else if (corrInv == "Y")
                        //oMatrix.Columns.Item("U_baseDocT").Cells.Item(oMatrix.RowCount).Specific.Value = "1";
                        oComboBox.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    else
                        //oMatrix.Columns.Item("U_baseDocT").Cells.Item(oMatrix.RowCount).Specific.Value = "0";
                        oComboBox.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oMatrix.Columns.Item("U_baseDoc").Cells.Item(oMatrix.RowCount).Specific.Value = 0;
                    oMatrix.Columns.Item("U_baseDTxt").Cells.Item(oMatrix.RowCount).Specific.Value = "";
                    oMatrix.Columns.Item("U_amtBsDc").Cells.Item(oMatrix.RowCount).Specific.Value = 0;
                    oMatrix.Columns.Item("U_tAmtBsDc").Cells.Item(oMatrix.RowCount).Specific.Value = 0;
                    oMatrix.Columns.Item("U_wbNumber").Cells.Item(oMatrix.RowCount).Specific.Value = "";
                    oMatrix.CommonSetting.SetRowBackColor(oMatrix.RowCount, -1);

                }

            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
                oForm.Freeze(false);
            }
        }

        public static void chooseInvoiceDoc(SAPbouiCOM.Form oDocForm, string docType, out string errorText)
        {
            //errorText = null;
            //oForm.Freeze(true);

            errorText = null;

            if (docType == "DownPaymentInvoice")
            {

                int left = 558 + 500;
                int Top = 300;

                //ფორმის აუცილებელი თვისებები
                Dictionary<string, object> formProperties = new Dictionary<string, object>();
                formProperties.Add("UniqueID", "BDO_TaxInvoiceReceivedChoose");
                formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
                formProperties.Add("Title", BDOSResources.getTranslate("TaxInvoiceReceived"));
                formProperties.Add("Left", left);
                formProperties.Add("Width", 700);
                formProperties.Add("Top", Top);
                formProperties.Add("Height", 200);
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

                        //ზედნადებების ცხრილი
                        string itemName = "InvDocs";
                        formItems = new Dictionary<string, object>();
                        formItems.Add("isDataSource", true);
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                        formItems.Add("Left", left);
                        formItems.Add("Width", oForm.Width - 20);
                        formItems.Add("Top", Top);
                        // formItems.Add("Height", 200);
                        formItems.Add("Height", oForm.Height - 20);
                        formItems.Add("UID", itemName);

                        formItems.Add("DisplayDesc", true);
                        formItems.Add("AffectsFormMode", false);
                        formItems.Add("State", SAPbouiCOM.BoFormStateEnum.fs_Maximized);

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }


                        SAPbouiCOM.DataTable oDataTable;
                        oDataTable = oForm.DataSources.DataTables.Add("InvDocs");

                        oDataTable.Columns.Add("#", SAPbouiCOM.BoFieldsType.ft_Text, 20); // 0 - ინდექსი გვჭირდება SetValue-ს პირველ პარემტრად                    
                        oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Text);
                        oDataTable.Columns.Add("cardCode", SAPbouiCOM.BoFieldsType.ft_Text, 15);
                        oDataTable.Columns.Add("cardCodeN", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                        oDataTable.Columns.Add("cardCodeT", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                        oDataTable.Columns.Add("series", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                        oDataTable.Columns.Add("number", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                        oDataTable.Columns.Add("invID", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                        oDataTable.Columns.Add("docDate", SAPbouiCOM.BoFieldsType.ft_Date, 20);
                        oDataTable.Columns.Add("opDate", SAPbouiCOM.BoFieldsType.ft_Date, 20);
                        oDataTable.Columns.Add("amount", SAPbouiCOM.BoFieldsType.ft_Sum, 20);
                        oDataTable.Columns.Add("VAT", SAPbouiCOM.BoFieldsType.ft_Sum, 20);
                        oDataTable.Columns.Add("OpenVAT", SAPbouiCOM.BoFieldsType.ft_Sum, 20);
                        oDataTable.Columns.Add("Remark", SAPbouiCOM.BoFieldsType.ft_Text);

                        string cardCode = oDocForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_cardCode", 0);
                        string docDate = oDocForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_docDate", 0);
                        string docEntry = oDocForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("DocEntry", 0);

                        string opDate = oDocForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_opDate", 0);

                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        string query = @"SELECT
                        ""BDO_TAXR"".""DocEntry"",
                        CASE WHEN ""BDO_TAXR"".""U_corrInv"" = 'Y' THEN ""BDO_TAXR"".""U_amtACor"" ELSE ""BDO_TAXR"".""U_amount"" END ""U_amount"",                          
                        CASE WHEN ""BDO_TAXR"".""U_corrInv"" = 'Y' THEN ""BDO_TAXR"".""U_amtTXACr"" ELSE ""BDO_TAXR"".""U_amountTX"" END ""U_amountTX"",
                        CASE WHEN ""BDO_TAXR"".""U_corrInv"" = 'Y' THEN ""BDO_TAXR"".""U_amtTXACr"" ELSE ""BDO_TAXR"".""U_amountTX"" END - IFNULL (""closedVatAmounts"".""closedVat"", 0) AS ""openVat"",
                        ""BDO_TAXR"".""U_cardCode"",
                        ""BDO_TAXR"".""U_cardCodeN"",
                        ""BDO_TAXR"".""U_cardCodeT"",
                        ""BDO_TAXR"".""U_series"",
                        ""BDO_TAXR"".""U_number"",
                        ""BDO_TAXR"".""U_invID"",
                        ""BDO_TAXR"".""U_docDate"",
                        ""BDO_TAXR"".""U_opDate"",
                        ""BDO_TAXR"".""Remark""
                        FROM ""@BDO_TAXR"" AS ""BDO_TAXR""
                                LEFT JOIN (
	                                SELECT 
                                SUM(""BDO_TXR5"".""U_drg_amount"") AS ""closedVat"",
	                                ""BDO_TXR5"".""DocEntry""
	                                 FROM ""@BDO_TXR5"" AS ""BDO_TXR5""
	                                 GROUP BY ""BDO_TXR5"".""DocEntry""
	                                ) AS ""closedVatAmounts""
	                                ON ""closedVatAmounts"".""DocEntry"" = ""BDO_TAXR"".""DocEntry""  

                        WHERE ""BDO_TAXR"".""U_downPaymnt"" = 'Y'
                        AND ""BDO_TAXR"".""Canceled"" = 'N'
                        AND (""BDO_TAXR"".""U_status"" = 'confirmed' OR ""BDO_TAXR"".""U_status"" = 'correctionConfirmed')
                        AND ""BDO_TAXR"".""U_cardCode"" = '" + cardCode + @"'
                        AND ""BDO_TAXR"".""DocEntry"" <> '" + docEntry + @"'                     
                        AND ""BDO_TAXR"".""U_opDate"" <= '" + opDate + @"'
                        AND CASE WHEN ""BDO_TAXR"".""U_corrInv"" = 'Y' THEN ""BDO_TAXR"".""U_amtTXACr"" ELSE ""BDO_TAXR"".""U_amountTX"" END - IFNULL (""closedVatAmounts"".""closedVat"", 0)  > 0
                        
                        ORDER bY ""BDO_TAXR"".""U_docDate"" DESC";


                        //LEFT JOIN (SELECT
                        //              ""DocEntry"",
                        //              ""U_corrDoc""
                        //              FROM ""@BDO_TAXR"" 
                        //              WHERE ""Canceled"" = 'N' AND ""U_corrInv"" = 'Y') AS ""Corr_TAXR"" 
                        //              ON ""BDO_TAXR"".""DocEntry"" = ""Corr_TAXR"".""U_corrDoc"" 

                        //AND ""Corr_TAXR"".""DocEntry"" IS NULL

                        if (Program.oCompany.DbServerType != SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            query = query.Replace("IFNULL", "ISNULL");
                        }

                        oRecordSet.DoQuery(query);

                        int rowIndex = 0;

                        while (!oRecordSet.EoF)
                        {

                            oDataTable.Rows.Add();
                            oDataTable.SetValue(0, rowIndex, rowIndex + 1);
                            oDataTable.SetValue(1, rowIndex, oRecordSet.Fields.Item("DocEntry").Value);
                            oDataTable.SetValue(2, rowIndex, oRecordSet.Fields.Item("U_cardCode").Value);
                            oDataTable.SetValue(3, rowIndex, oRecordSet.Fields.Item("U_cardCodeN").Value);
                            oDataTable.SetValue(4, rowIndex, oRecordSet.Fields.Item("U_cardCodeT").Value);
                            oDataTable.SetValue(5, rowIndex, oRecordSet.Fields.Item("U_series").Value);
                            oDataTable.SetValue(6, rowIndex, oRecordSet.Fields.Item("U_number").Value);
                            oDataTable.SetValue(7, rowIndex, oRecordSet.Fields.Item("U_invID").Value);
                            oDataTable.SetValue(8, rowIndex, oRecordSet.Fields.Item("U_docDate").Value);
                            oDataTable.SetValue(9, rowIndex, oRecordSet.Fields.Item("U_opDate").Value);
                            oDataTable.SetValue(10, rowIndex, oRecordSet.Fields.Item("U_amount").Value);
                            oDataTable.SetValue(11, rowIndex, oRecordSet.Fields.Item("U_amountTX").Value);
                            oDataTable.SetValue(12, rowIndex, oRecordSet.Fields.Item("openVat").Value);
                            oDataTable.SetValue(13, rowIndex, oRecordSet.Fields.Item("Remark").Value);

                            //DateTime docDate = new DateTime(1, 1, 1);

                            //if (DateTime.TryParseExact(oRecordSet.Fields.Item("U_docDate").Value.Replace("T", " "), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out docDate) == false)
                            //{
                            //    docDate = new DateTime(1, 1, 1);
                            //}

                            //oDataTable.SetValue(8, rowIndex, docDate);                        

                            rowIndex++;

                            oRecordSet.MoveNext();

                            //if (rowIndex==10) break;

                            //break;
                        }

                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvDocs").Specific));
                        SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                        SAPbouiCOM.Column oColumn;

                        int columnWidth = (oForm.Width - 20) / 12;

                        oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = "#";
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "#");
                        oColumn.TitleObject.Sortable = true;

                        oColumn = oColumns.Add("DocEntry", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocEntry");
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "DocEntry");
                        oColumn.TitleObject.Sortable = true;

                        oColumn = oColumns.Add("cardCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("SupplierCode");
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "cardCode");

                        oColumn = oColumns.Add("cardCodeN", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("SupplierName");
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "cardCodeN");

                        oColumn = oColumns.Add("cardCodeT", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("SupplierTIN");
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "cardCodeT");

                        oColumn = oColumns.Add("series", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("SeriesNumber");
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "series");
                        oColumn.TitleObject.Sortable = true;

                        oColumn = oColumns.Add("number", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("InvoiceNumber");
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "number");
                        oColumn.TitleObject.Sortable = true;

                        oColumn = oColumns.Add("invID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("TaxInvoiceID");
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "invID");
                        oColumn.TitleObject.Sortable = true;

                        oColumn = oColumns.Add("docDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("PostingDate");
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "docDate");
                        oColumn.TitleObject.Sortable = true;

                        oColumn = oColumns.Add("opDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("OperationMonth");
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "opDate");
                        oColumn.TitleObject.Sortable = true;

                        oColumn = oColumns.Add("amount", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("Amount");
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "amount");
                        oColumn.TitleObject.Sortable = true;

                        oColumn = oColumns.Add("VAT", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatAmount");
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "VAT");
                        oColumn.TitleObject.Sortable = true;

                        oColumn = oColumns.Add("OpenVAT", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("OpenVAT");
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "OpenVAT");
                        oColumn.TitleObject.Sortable = true;

                        oColumn = oColumns.Add("Remark", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("Remarks");
                        oColumn.Width = columnWidth;
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind("InvDocs", "Remark");
                        oColumn.TitleObject.Sortable = true;

                        oMatrix.Clear();
                        oMatrix.LoadFromDataSource();
                        oMatrix.AutoResizeColumns();

                        ////Choose ღილაკი
                        //formItems = new Dictionary<string, object>();
                        //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                        //formItems.Add("Left", 13);
                        //formItems.Add("Width", 100);
                        //formItems.Add("Top", oForm.Height - 10);
                        //formItems.Add("Height", 10);
                        //formItems.Add("Caption", BDOSResources.getTranslate("Choose"));
                        //formItems.Add("UID", "ChsInvDocs");

                        //FormsB1.createFormItem(oForm, formItems, out errorText);
                        //if (errorText != null)
                        //{
                        //    return;
                        //}
                    }



                    oForm.Visible = true;
                    //oForm.Select();

                }


            }

        }

        public static void delMatrixRow(SAPbouiCOM.Form oForm, string ItemUID, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            try
            {

                if (ItemUID == "delMTRB")
                {

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));
                    oMatrix.FlushToDataSource();
                    int firstRow = 0;
                    int row = 0;
                    int deletedRowCount = 0;

                    while (row != -1)
                    {
                        row = oMatrix.GetNextSelectedRow(firstRow, SAPbouiCOM.BoOrderType.ot_RowOrder);
                        if (row > -1)
                        {
                            deletedRowCount++;
                            oForm.DataSources.DBDataSources.Item("@BDO_TXR1").RemoveRecord(row - deletedRowCount);
                            firstRow = row;
                        }
                    }

                    oMatrix.LoadFromDataSource();
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && deletedRowCount > 0)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }

                }

                else if (ItemUID == "delDPinv")
                {
                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("DPinvoices").Specific));
                    oMatrix.FlushToDataSource();
                    int firstRow = 0;
                    int row = 0;
                    int deletedRowCount = 0;

                    while (row != -1)
                    {
                        row = oMatrix.GetNextSelectedRow(firstRow, SAPbouiCOM.BoOrderType.ot_RowOrder);
                        if (row > -1)
                        {
                            deletedRowCount++;
                            oForm.DataSources.DBDataSources.Item("@BDO_TXR4").RemoveRecord(row - deletedRowCount);
                            firstRow = row;
                        }
                    }

                    oMatrix.LoadFromDataSource();
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && deletedRowCount > 0)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }

                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
                oForm.Freeze(false);
            }
        }

        private static string getLinkStatus(decimal amount, decimal amountTX, decimal amtBsDc, decimal tAmtBsDc)
        {
            string linkStatus = null;

            if (amtBsDc == amount & tAmtBsDc == amountTX)
            {
                linkStatus = "1"; //მიბმული                        
            }
            else if (amtBsDc > 0)
            {
                linkStatus = "2"; //ნაწილობრივ მიბმული                        
            }
            else
            {
                linkStatus = "0"; //მიუბმელი                      
            }
            return linkStatus;
        }

        public static void formDataAddUpdate(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                //string statusInv = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_status", 0).Trim();
                //if (statusInv == "confirmed" || statusInv == "correctionConfirmed")
                //{
                //    errorText = BDOSResources.getTranslate("UpdateConfirmedTaxInvoiceNotAllowed");
                //    return;
                //}

                //
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDO_TXR1");

                Dictionary<string, SAPbouiCOM.DBDataSource> oKeysDictionary = new Dictionary<string, SAPbouiCOM.DBDataSource>();
                oKeysDictionary.Add("U_baseDocT", oDBDataSource);
                oKeysDictionary.Add("U_baseDoc", oDBDataSource);
                oKeysDictionary.Add("U_baseDTxt", oDBDataSource);
                oKeysDictionary.Add("U_wbNumber", oDBDataSource);

                CommonFunctions.checkDuplicatesInDBDataSources(oDBDataSource, oKeysDictionary, out errorText);
                if (string.IsNullOrEmpty(errorText) == false)
                {
                    return;
                }
                //

                decimal amount = Convert.ToDecimal(oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_amount", 0), CultureInfo.InvariantCulture); //თანხა დღგ-ის ჩათვლით
                decimal amountTX = Convert.ToDecimal(oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_amountTX", 0), CultureInfo.InvariantCulture); //დღგ-ის თანხა
                decimal amtBsDc = 0; //თანხა დღგ-ის ჩათვლით (საფუძველი დოკუმენტის)
                decimal tAmtBsDc = 0; //დღგ-ის თანხა (საფუძველი დოკუმენტის)

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));
                SAPbouiCOM.EditText oEditText;

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    oEditText = oMatrix.Columns.Item("U_amtBsDc").Cells.Item(i).Specific;
                    amtBsDc = amtBsDc + Convert.ToDecimal(oEditText.Value, CultureInfo.InvariantCulture);
                    oEditText = oMatrix.Columns.Item("U_tAmtBsDc").Cells.Item(i).Specific;
                    tAmtBsDc = tAmtBsDc + Convert.ToDecimal(oEditText.Value, CultureInfo.InvariantCulture);
                }

                string linkStatus = getLinkStatus(amount, amountTX, amtBsDc, tAmtBsDc);

                oForm.DataSources.DBDataSources.Item("@BDO_TAXR").SetValue("U_LinkStatus", 0, linkStatus);
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

        private static void getInfoDoc(string invID, string number, out int docEntry, out decimal amtACor, out decimal amtTXACr, out List<string> wbNumber, out string errorText)
        {
            errorText = null;
            docEntry = 0;
            amtACor = 0;
            amtTXACr = 0;
            wbNumber = new List<string>();
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = @"SELECT
                ""BDO_TAXR"".""DocEntry"",
                CASE WHEN ""BDO_TAXR"".""U_corrInv"" = 'Y' THEN ""BDO_TAXR"".""U_amtACor"" ELSE ""BDO_TAXR"".""U_amount"" END AS ""amtACor"",
                CASE WHEN ""BDO_TAXR"".""U_corrInv"" = 'Y' THEN ""BDO_TAXR"".""U_amtTXACr"" ELSE ""BDO_TAXR"".""U_amountTX"" END AS ""amtTXACr"",
                ""BDO_TAXR"".""U_invID"",
                ""BDO_TXR2"".""U_wbNumber""
                FROM ""@BDO_TAXR"" AS ""BDO_TAXR""
                LEFT JOIN ""@BDO_TXR2"" AS ""BDO_TXR2"" 
                ON ""BDO_TAXR"".""DocEntry"" = ""BDO_TXR2"".""DocEntry""
                WHERE ""BDO_TAXR"".""Canceled"" = 'N'"; // AND ""BDO_TAXR"".""U_status"" NOT IN ('removed', 'canceled')

                if (invID != null)
                {
                    query = query + @"AND ""BDO_TAXR"".""U_invID"" = '" + invID + "'";
                }
                else
                {
                    query = query + @"AND ""BDO_TAXR"".""U_number"" = '" + number + "'";
                }

                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    docEntry = oRecordSet.Fields.Item("DocEntry").Value;
                    amtACor = Convert.ToDecimal(oRecordSet.Fields.Item("amtACor").Value);
                    amtTXACr = Convert.ToDecimal(oRecordSet.Fields.Item("amtTXACr").Value);
                    wbNumber.Add(oRecordSet.Fields.Item("U_wbNumber").Value);
                    oRecordSet.MoveNext();
                    //break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
        }

        public static void createDocumentTaxInvoiceType(TaxInvoice oTaxInvoice, bool isUpdate, DataRow taxDataRow, out string errorText)
        {
            errorText = null;

            int docEntry = 0;
            decimal amtCor = 0;
            decimal amtTXCor = 0;
            string statusRS = null;
            string number = taxDataRow["F_NUMBER"].ToString(); //ა/ფ ნომერი
            string ID = taxDataRow["ID"].ToString(); //ა/ფ ID
            List<string> wbNumberCor = new List<string>();

            getInfoDoc(null, number, out docEntry, out amtCor, out amtTXCor, out wbNumberCor, out errorText);

            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            oCompanyService = Program.oCompany.GetCompanyService();
            oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXR_D");
            oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

            if (docEntry == 0)
            {
                var response = oGeneralService.Add(oGeneralData);
                docEntry = response.GetProperty("DocEntry");

                operationRS(oTaxInvoice, "create", docEntry, -1, new DateTime(), taxDataRow, out statusRS, out errorText);
            }
            else if (isUpdate == true)
            {
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", docEntry);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                string status = oGeneralData.GetProperty("U_status");

                if (status != "confirmed" && status != "correctionConfirmed")
                {
                    operationRS(oTaxInvoice, "update", docEntry, -1, new DateTime(), taxDataRow, out statusRS, out errorText);
                }
            }
        }

        public static void chooseFromListForBaseDocs(SAPbouiCOM.Form oForm, string taxDocEntryStr, out string errorText)
        {
            errorText = null;

            int taxDocEntry = string.IsNullOrEmpty(taxDocEntryStr) ? 0 : Convert.ToInt32(taxDocEntryStr);

            string caption = BDOSResources.getTranslate("ChooseTaxInvoice");
            string taxID = "";
            string taxNumber = "";
            string taxSeries = "";
            string taxStatus = "";
            string taxCreateDate = "";

            if (taxDocEntry != 0)
            {
                Dictionary<string, object> taxDocInfo = getTaxInvoiceReceivedDocumentInfo(taxDocEntry, out errorText);
                if (taxDocInfo != null)
                {
                    taxDocEntry = Convert.ToInt32(taxDocInfo["docEntry"]);
                    taxID = taxDocInfo["invID"].ToString();
                    taxNumber = taxDocInfo["number"].ToString();
                    taxSeries = taxDocInfo["series"].ToString();
                    taxStatus = taxDocInfo["status"].ToString();
                    taxCreateDate = taxDocInfo["createDate"].ToString();

                    if (taxDocEntry != 0)
                    {
                        DateTime taxCreateDateDT = DateTime.ParseExact(taxCreateDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                        if (taxSeries == "")
                        {
                            caption = BDOSResources.getTranslate("TaxInvoiceDate") + " " + taxCreateDateDT;
                        }
                        else
                        {
                            caption = BDOSResources.getTranslate("TaxInvoiceSeries") + " " + taxSeries + " № " + taxNumber + " " + BDOSResources.getTranslate("Data") + " " + taxCreateDateDT;
                        }
                    }
                }
            }
            else
            {
                taxDocEntry = 0;
            }

            oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx = taxDocEntry == 0 ? "" : taxDocEntry.ToString();
            oForm.DataSources.UserDataSources.Item("BDO_TaxSer").ValueEx = taxSeries;
            oForm.DataSources.UserDataSources.Item("BDO_TaxNum").ValueEx = taxNumber;
            oForm.DataSources.UserDataSources.Item("BDO_TaxDat").ValueEx = taxCreateDate;

            SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_TaxTxt").Specific;
            oStaticText.Caption = caption;

            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
        }

        public static void receiveVAT(int docEntry, DateTime declDate, string YesNoEmpty, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            oCompanyService = Program.oCompany.GetCompanyService();
            oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXR_D");

            SAPbobsCOM.GeneralDataParams UDOParameter = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            UDOParameter.SetProperty("DocEntry", docEntry);

            oGeneralData = oGeneralService.GetByParams(UDOParameter);

            string U_recStatus = oGeneralData.GetProperty("U_TaxInRcvd");


            try
            {
                if (YesNoEmpty == "Y")
                {
                    if (U_recStatus == "Y")
                    {
                        errorText = BDOSResources.getTranslate("TaxInvoiceAlreadyReceived");
                        return;
                    }
                    oGeneralData.SetProperty("U_vatRDate", declDate);
                    oGeneralData.SetProperty("U_TaxInRcvd", "Y");
                    oGeneralService.Update(oGeneralData);
                }

                else if (YesNoEmpty == "N")
                {
                    oGeneralData.SetProperty("U_vatRDate", "");
                    oGeneralData.SetProperty("U_TaxInRcvd", "N");
                    oGeneralService.Update(oGeneralData);
                }
                else
                {
                    oGeneralData.SetProperty("U_vatRDate", "");
                    oGeneralData.SetProperty("U_TaxInRcvd", "");
                    oGeneralService.Update(oGeneralData);
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
                GC.Collect();
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == true)
            {
                return;
            }

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    if (oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_docDate", 0) == "")
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("PostingDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                        Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                        BubbleEvent = false;
                    }
                }
            }

            if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_TAXR_D")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
                {
                    formDataLoad(oForm, out errorText);
                    setVisibleFormItems(oForm, out errorText);
                }

                if ((BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) && BusinessObjectInfo.BeforeAction)
                {
                    if (Program.cancellationTrans && Program.canceledDocEntry != 0)
                    {
                        SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDO_TAXR");

                        string docEntry = oDBDataSource.GetValue("DocEntry", 0).Trim();

                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = @"SELECT ""BDO_TAXR"".""DocEntry"",
                            ""BDO_TXR4"".""U_DP_invoice""
                            FROM ""@BDO_TAXR"" AS ""BDO_TAXR""
                            LEFT JOIN ""@BDO_TXR4"" AS ""BDO_TXR4""
                            ON ""BDO_TAXR"".""DocEntry"" = ""BDO_TXR4"".""DocEntry""
                            WHERE ""BDO_TAXR"".""Canceled"" = 'N'
                            AND (""BDO_TXR4"".""U_DP_invoice"" =  '" + docEntry + @"' OR ""BDO_TAXR"".""U_corrDoc"" = '" + docEntry + "')";

                        oRecordSet.DoQuery(query);

                        if (oRecordSet.RecordCount > 0)
                        {
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("CanNotCancelLinkedDocument"));
                            BubbleEvent = false;
                        }
                    }
                    else
                    {
                        formDataAddUpdate(oForm, out errorText);
                        if (string.IsNullOrEmpty(errorText) == false)
                        {
                            Program.uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                        }
                    }
                }

                if ((BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
                    && BusinessObjectInfo.ActionSuccess != BusinessObjectInfo.BeforeAction
                    && Program.canceledDocEntry == 0)
                {
                    SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(0);
                    string invoiceStatus = oDBDataSource.GetValue("U_status", 0).Trim();
                    string docEntry = oDBDataSource.GetValue("DocEntry", 0);
                    string docStatus = oForm.Items.Item("StatusC").Specific.Value;
                    string jrnEnt = oDBDataSource.GetValue("U_JrnEnt", 0);
                    string downPaymnt = oDBDataSource.GetValue("U_downPaymnt", 0).Trim();
                    string corrInv = oDBDataSource.GetValue("U_corrInv", 0).Trim();
                    decimal docVatAmount = Convert.ToDecimal(oDBDataSource.GetValue("U_amountTX", 0), CultureInfo.InvariantCulture);

                    SAPbouiCOM.DBDataSource DBDataSourceTable = oForm.DataSources.DBDataSources.Item("@BDO_TXR4");

                    decimal vatAmountSum = 0;
                    decimal drg_amount = 0;
                    decimal openVat = 0;

                    int JEcount = DBDataSourceTable.Size;

                    for (int i = 0; i < JEcount; i++)
                    {
                        drg_amount = Convert.ToDecimal(DBDataSourceTable.GetValue("U_drg_amount", i), CultureInfo.InvariantCulture);
                        openVat = Convert.ToDecimal(DBDataSourceTable.GetValue("U_max_amount", i), CultureInfo.InvariantCulture);

                        if (drg_amount > openVat)
                        {
                            int row = i + 1;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("VatAmountCanNotBeMoreThanOpenVatRow") + row);
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                            BubbleEvent = false;
                            break;
                        }
                        vatAmountSum = vatAmountSum + drg_amount;
                    }

                    if (BubbleEvent == true && vatAmountSum > docVatAmount)
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("AdvancesVatAmountCanNotBeMoreThanTaxInvoiceVatAmount"));
                        Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                        BubbleEvent = false;
                    }

                    else if (downPaymnt != "Y" && BubbleEvent && docStatus == "O"
                        && (invoiceStatus == "confirmed" || (invoiceStatus == "corrected" && corrInv != "Y") || invoiceStatus == "correctionConfirmed")
                        && (string.IsNullOrEmpty(jrnEnt) || jrnEnt == "0"))
                    {
                        string DocNum = oDBDataSource.GetValue("DocNum", 0);

                        DateTime DocDate = DateTime.ParseExact(oDBDataSource.GetValue("U_docDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                        CommonFunctions.StartTransaction();

                        Program.JrnLinesGlobal = new DataTable();
                        DataTable JrnLinesDT = createAdditionalEntries(oForm, null);

                        JrnEntry(docEntry, DocNum, DocDate, JrnLinesDT, out errorText);
                        if (errorText != null)
                        {
                            Program.uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                        }
                        else
                        {
                            if (!BusinessObjectInfo.ActionSuccess)
                            {
                                Program.JrnLinesGlobal = JrnLinesDT;
                            }
                            else
                            {
                                // გატარებები
                                string Ref1 = docEntry;
                                string Ref2 = "UDO_F_BDO_TAXR_D";

                                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string query = "SELECT " +
                                                "*  " +
                                                "FROM \"OJDT\"  " +
                                                "WHERE \"StornoToTr\" IS NULL " +
                                                "AND \"Ref1\" = '" + Ref1 + "' " +
                                                "AND \"Ref2\" = '" + Ref2 + "' ";
                                oRecordSet.DoQuery(query);
                                int U_JrnEnt = 0;
                                if (!oRecordSet.EoF)
                                {
                                    U_JrnEnt = oRecordSet.Fields.Item("TransId").Value;
                                    oDBDataSource.SetValue("U_JrnEnt", 0, Convert.ToString(oRecordSet.Fields.Item("TransId").Value, CultureInfo.InvariantCulture));
                                }
                                //else                                
                                //{
                                //    oDBDataSource.SetValue("U_JrnEnt", 0, "");
                                //}                                    
                                if (U_JrnEnt != 0)
                                {
                                    SAPbobsCOM.CompanyService oCompanyService = null;
                                    SAPbobsCOM.GeneralService oGeneralService = null;
                                    SAPbobsCOM.GeneralData oGeneralDataInv = null;
                                    SAPbobsCOM.GeneralDataParams oGeneralParams = null;

                                    oCompanyService = Program.oCompany.GetCompanyService();
                                    oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXR_D");
                                    oGeneralDataInv = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                                    oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                                    oGeneralParams.SetProperty("DocEntry", docEntry);
                                    oGeneralDataInv = oGeneralService.GetByParams(oGeneralParams);
                                    oGeneralDataInv.SetProperty("U_JrnEnt", Convert.ToString(U_JrnEnt, CultureInfo.InvariantCulture));

                                    oGeneralService.Update(oGeneralDataInv);

                                    Marshal.FinalReleaseComObject(oGeneralService);
                                }
                            }
                        }

                        if (Program.oCompany.InTransaction)
                        {
                            //თუ დოკუმენტი გატარდა, მერე ვაკეთებს ბუღალტრულ გატარებას
                            if (BusinessObjectInfo.ActionSuccess == true & BusinessObjectInfo.BeforeAction == false)
                            {
                                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            }
                            else
                            {
                                oDBDataSource.SetValue("U_JrnEnt", 0, "");
                                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            }
                        }
                        else
                        {
                            oDBDataSource.SetValue("U_JrnEnt", 0, "");

                            Program.uiApp.MessageBox("ტრანზაქციის დასრულებს შეცდომა");
                            BubbleEvent = false;
                        }
                    }
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE && !BusinessObjectInfo.BeforeAction && BusinessObjectInfo.ActionSuccess)
                {
                    if (Program.cancellationTrans && Program.canceledDocEntry != 0)
                    {
                        cancellation(oForm, Program.canceledDocEntry, out errorText);
                        Program.canceledDocEntry = 0;
                    }
                }
            }
        }

        public static void cancellation(SAPbouiCOM.Form oForm, int docEntry, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.cancellation(oForm, docEntry, "UDO_F_BDO_TAXR_D", out errorText);

                if (errorText == null)
                {

                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string query = @"DELETE
                        FROM ""@BDO_TXR5"" 
                        WHERE ""@BDO_TXR5"".""U_tax_invoice"" = '" + docEntry + @"'";

                    oRecordSet.DoQuery(query);

                    while (!oRecordSet.EoF)
                    {
                    }

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
                GC.Collect();
            }
        }

        public static DataTable createAdditionalEntries(SAPbouiCOM.Form oForm, SAPbobsCOM.GeneralData oGeneralData)
        {
            DataTable jeLines = JournalEntry.JournalEntryTable();

            SAPbobsCOM.GeneralDataCollection oChild = null;
            SAPbouiCOM.DBDataSource DBDataSourceTable = null;
            SAPbouiCOM.DBDataSource oDBDataSource = null;
            bool downPaymnt;
            bool correctionConfirmed;
            string corrDocNum = "";
            string DocEntry;
            DateTime DocDate;
            DateTime? taxDate = null;
            DataTable AccountTable = CommonFunctions.GetOACTTable();
            int JEcount = 0;

            if (oForm == null)
            {
                downPaymnt = oGeneralData.GetProperty("U_downPaymnt") == "Y";
                correctionConfirmed = oGeneralData.GetProperty("U_status").Trim() == "correctionConfirmed";

                if (oGeneralData.GetProperty("U_corrDoc") > 0)
                {
                    corrDocNum = Convert.ToString(oGeneralData.GetProperty("U_corrDoc"), CultureInfo.InvariantCulture);
                }

                if (downPaymnt)
                {
                    oChild = oGeneralData.Child("BDO_TXR3");
                }
                else
                {
                    oChild = oGeneralData.Child("BDO_TXR4");
                }

                JEcount = oChild.Count;

                DocDate = oGeneralData.GetProperty("U_docDate");
                DocEntry = Convert.ToString(oGeneralData.GetProperty("DocEntry"), CultureInfo.InvariantCulture);
            }
            else
            {
                oDBDataSource = oForm.DataSources.DBDataSources.Item(0);
                downPaymnt = oDBDataSource.GetValue("U_downPaymnt", 0) == "Y";
                correctionConfirmed = oDBDataSource.GetValue("U_status", 0).Trim() == "correctionConfirmed";

                corrDocNum = oDBDataSource.GetValue("U_corrDoc", 0);

                if (downPaymnt)
                {
                    DBDataSourceTable = oForm.DataSources.DBDataSources.Item("@BDO_TXR3");
                }
                else
                {
                    DBDataSourceTable = oForm.DataSources.DBDataSources.Item("@BDO_TXR4");
                }

                JEcount = DBDataSourceTable.Size;

                DocDate = DateTime.ParseExact(oDBDataSource.GetValue("U_docDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
                DocEntry = oDBDataSource.GetValue("DocEntry", 0);
            }

            if (JEcount == 0) return jeLines;

            if (downPaymnt)
            {
                if (oForm == null)
                    taxDate = oGeneralData.GetProperty("U_taxDate");
                else
                    taxDate = DateTime.ParseExact(oDBDataSource.GetValue("U_taxDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                SAPbobsCOM.Recordset oRecordSet;
                string query = "";

                DataTable correctedDataTable = new DataTable();
                correctedDataTable.Columns.Add("DebitAccount", typeof(string));
                correctedDataTable.Columns.Add("vatAmount", typeof(decimal));

                decimal vatAmount = 0;

                DataTable newDataTable = correctedDataTable.Clone();

                if (corrDocNum != "")
                {
                    oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    query = @"SELECT 
                            ""vatAccounts"".""Account"",
                            0 - SUM(""BDO_TXR3"".""U_drg_amount"") AS ""U_drg_amount""
                            FROM ""@BDO_TXR3"" AS ""BDO_TXR3""
                            LEFT JOIN (
	                            SELECT
	                            ""OVTG"".""Code"",
	                            ""OVTG"".""Account""
	                            FROM ""OVTG""
	                            ) AS ""vatAccounts""
                            ON ""vatAccounts"".""Code"" = ""BDO_TXR3"".""U_vat_type""
                            WHERE ""BDO_TXR3"".""DocEntry"" = '" + corrDocNum + @"'
                            GROUP BY ""vatAccounts"".""Account""";

                    oRecordSet.DoQuery(query);

                    int recordCount = oRecordSet.RecordCount;

                    while (!oRecordSet.EoF)
                    {

                        //vatAmount = oRecordSet.Fields.Item("U_drg_amount").Value;

                        DataRow DataRow = correctedDataTable.Rows.Add();
                        DataRow["DebitAccount"] = oRecordSet.Fields.Item("Account").Value;
                        DataRow["vatAmount"] = oRecordSet.Fields.Item("U_drg_amount").Value;

                        oRecordSet.MoveNext();
                    }
                }

                string year = taxDate == null ? DocDate.Year.ToString() : taxDate?.Year.ToString();
                string DebitAccount = "";
                //string CreditAccount = CommonFunctions.getOADM( "U_BDO_TaxAcc").ToString();
                string CreditAccount = CommonFunctions.getPeriodsCategory("PurcVatOff", year);

                oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                query = "";

                for (int i = 0; i < JEcount; i++)
                {

                    vatAmount = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, oChild, null, "U_drg_amount", i), CultureInfo.InvariantCulture);
                    string vatGrp = Convert.ToString(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, oChild, null, "U_vat_type", i), CultureInfo.InvariantCulture);

                    query = "SELECT " +
                                "* " +
                                "FROM \"OVTG\" " +
                                "WHERE \"OVTG\".\"Code\"='" + vatGrp + "'";

                    oRecordSet.DoQuery(query);

                    if (!oRecordSet.EoF)
                    {
                        DebitAccount = oRecordSet.Fields.Item("Account").Value;
                    }

                    DataRow DataRow = newDataTable.Rows.Add();
                    DataRow["DebitAccount"] = DebitAccount;
                    DataRow["vatAmount"] = vatAmount;

                }

                //DataTable finalDataTable = newDataTable.AsEnumerable().Union(correctedDataTable.AsEnumerable()).Distinct(DataRowComparer.Default).CopyToDataTable<DataRow>();
                DataTable finalDataTable = newDataTable.AsEnumerable().Union(correctedDataTable.AsEnumerable()).Distinct().CopyToDataTable<DataRow>();

                DataTable finalTable = finalDataTable.AsEnumerable().GroupBy(row => new
                {
                    DebitAccount = row.Field<string>("DebitAccount")

                }).Select(g =>
                {
                    var row = finalDataTable.NewRow();
                    row["DebitAccount"] = g.Key.DebitAccount;
                    row["vatAmount"] = g.Sum(r => r.Field<decimal>("vatAmount"));
                    return row;
                    //}).CopyToDataTable<DataRow>();
                }).CopyToDataTable();

                for (int i = 0; i < finalTable.Rows.Count; i++)
                {
                    DataRow dtRow = finalTable.Rows[i];
                    DebitAccount = dtRow["DebitAccount"].ToString();
                    vatAmount = Convert.ToDecimal(dtRow["vatAmount"]);

                    if (vatAmount != 0)
                    {
                        JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, vatAmount, 0, "", "", "", "", "", "", "", "", "");
                    }
                }

            }
            else
            {
                DataTable newDataTable = new DataTable();
                newDataTable.Columns.Add("CreditAccount", typeof(string));
                newDataTable.Columns.Add("DebitAccount", typeof(string));
                newDataTable.Columns.Add("vatAmount", typeof(decimal));

                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = "";

                string year = DocDate.Year.ToString();

                string CreditAccount = "";
                //string CreditAccount = CommonFunctions.getOADM( "U_BDO_TaxAcc").ToString();
                string DebitAccount = CommonFunctions.getPeriodsCategory("PurcVatOff", year);

                for (int i = 0; i < JEcount; i++)
                {
                    string invoiceDocEntry = Convert.ToString(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, oChild, null, "U_DP_invoice", i), CultureInfo.InvariantCulture);
                    if (!string.IsNullOrEmpty(invoiceDocEntry))
                    {
                        decimal vatAmount = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, oChild, null, "U_drg_amount", i), CultureInfo.InvariantCulture);
                        decimal currentVatAmount = vatAmount;
                        decimal openVat = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, oChild, null, "U_max_amount", i), CultureInfo.InvariantCulture);

                        SAPbobsCOM.CompanyService oCompanyService = null;
                        SAPbobsCOM.GeneralService oGeneralService = null;
                        SAPbobsCOM.GeneralData oGeneralDataInv = null;
                        SAPbobsCOM.GeneralDataParams oGeneralParams = null;

                        oCompanyService = Program.oCompany.GetCompanyService();
                        oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXR_D");
                        oGeneralDataInv = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        oGeneralParams.SetProperty("DocEntry", invoiceDocEntry);
                        oGeneralDataInv = oGeneralService.GetByParams(oGeneralParams);

                        decimal totalVatAmount = 0;

                        if (correctionConfirmed)
                        {
                            totalVatAmount = Convert.ToDecimal(oGeneralDataInv.GetProperty("U_amtTXCor"), CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            totalVatAmount = Convert.ToDecimal(oGeneralDataInv.GetProperty("U_amountTX"), CultureInfo.InvariantCulture);
                        }

                        decimal closedVat = totalVatAmount - openVat;

                        if (closedVat < 0)
                        {
                            closedVat = 0;
                        }

                        query = @"SELECT 
                            ""vatAccounts"".""Account"",
                            SUM(""BDO_TXR3"".""U_drg_amount"") AS ""U_drg_amount""
                            FROM ""@BDO_TXR3"" AS ""BDO_TXR3""
                            LEFT JOIN (
	                            SELECT
	                            ""OVTG"".""Code"",
	                            ""OVTG"".""Account""
	                            FROM ""OVTG""
	                            ) AS ""vatAccounts""
                            ON ""vatAccounts"".""Code"" = ""BDO_TXR3"".""U_vat_type""
                            WHERE ""BDO_TXR3"".""DocEntry"" = '" + invoiceDocEntry + @"'
                            GROUP BY ""vatAccounts"".""Account""";

                        oRecordSet.DoQuery(query);

                        int recordCount = oRecordSet.RecordCount;

                        while (!oRecordSet.EoF)
                        {
                            decimal rowVatAmount = Convert.ToDecimal(oRecordSet.Fields.Item("U_drg_amount").Value, CultureInfo.InvariantCulture);

                            if (closedVat - rowVatAmount > 0)
                            {
                                closedVat = closedVat - rowVatAmount;
                            }
                            else
                            {
                                decimal rowOpenVat = rowVatAmount - closedVat;

                                closedVat = 0;

                                if (rowOpenVat > 0 && vatAmount > 0)
                                {
                                    CreditAccount = oRecordSet.Fields.Item("Account").Value;

                                    DataRow DataRow = newDataTable.Rows.Add();
                                    DataRow["DebitAccount"] = DebitAccount;
                                    DataRow["CreditAccount"] = CreditAccount;

                                    rowVatAmount = rowOpenVat > vatAmount ? vatAmount : rowOpenVat;
                                    DataRow["vatAmount"] = rowVatAmount;
                                    vatAmount = vatAmount - rowVatAmount;
                                }
                            }
                            oRecordSet.MoveNext();
                        }

                        SAPbobsCOM.GeneralDataCollection oChildInv = oGeneralDataInv.Child("BDO_TXR5");
                        SAPbobsCOM.GeneralData oChildGeneralData = oChildInv.Add();

                        try
                        {
                            oChildGeneralData.SetProperty("U_tax_invoice", DocEntry);
                            oChildGeneralData.SetProperty("U_drg_amount", Convert.ToDouble(currentVatAmount));
                        }
                        catch { }

                        oGeneralService.Update(oGeneralDataInv);
                    }
                }

                if (newDataTable.Rows.Count > 0)
                {
                    DataTable finalTable = newDataTable.AsEnumerable().GroupBy(row => new
                    {
                        DebitAccount = row.Field<string>("DebitAccount"),
                        CreditAccount = row.Field<string>("CreditAccount")

                    }).Select(g =>
                    {
                        var row = newDataTable.NewRow();
                        row["DebitAccount"] = g.Key.DebitAccount;
                        row["CreditAccount"] = g.Key.CreditAccount;
                        row["vatAmount"] = g.Sum(r => r.Field<decimal>("vatAmount"));
                        return row;
                        //}).CopyToDataTable<DataRow>();
                    }).CopyToDataTable();

                    for (int i = 0; i < finalTable.Rows.Count; i++)
                    {
                        DataRow dtRow = finalTable.Rows[i];
                        DebitAccount = dtRow["DebitAccount"].ToString();
                        CreditAccount = dtRow["CreditAccount"].ToString();
                        decimal vatAmount = Convert.ToDecimal(dtRow["vatAmount"]);

                        if (vatAmount != 0)
                        {
                            JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, vatAmount, 0, "", "", "", "", "", "", "", "", "");
                        }
                    }
                }
            }
            return jeLines;
        }

        public static void JrnEntry(string DocEntry, string DocNum, DateTime DocDate, DataTable JrnLinesDT, out string errorText)
        {
            try
            {
                if (JrnLinesDT.Rows.Count == 0)
                {
                    errorText = BDOSResources.getTranslate("DataForJournalEntryNotFound");
                    return;
                }

                JournalEntry.JrnEntry(DocEntry, "UDO_F_BDO_TAXR_D", "Tax Invoice Received: " + DocNum, DocDate, JrnLinesDT, out errorText);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (FormUID == "BDO_TaxInvoiceReceivedChoose")
            {
                if (pVal.ItemUID == "InvDocs" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK && pVal.BeforeAction == false)
                {

                    SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    int row = pVal.Row;

                    if (row == 0) return;


                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvDocs").Specific;

                    string docEntry = oMatrix.Columns.Item("DocEntry").Cells.Item(row).Specific.Value;
                    string maxAmount = oMatrix.Columns.Item("OpenVAT").Cells.Item(row).Specific.Value;

                    SAPbouiCOM.Form oDocForm = Program.uiApp.Forms.GetForm("UDO_FT_UDO_F_BDO_TAXR_D", 1);
                    //SAPbouiCOM.DBDataSource DBDataSourceTable = oDocForm.DataSources.DBDataSources.Item("@BDO_TXR4");

                    oMatrix = ((SAPbouiCOM.Matrix)(oDocForm.Items.Item("DPinvoices").Specific));

                    int JEcount = oMatrix.RowCount;

                    for (int i = 0; i < JEcount; i++)
                    {
                        string invoiceDocEntry = oMatrix.Columns.Item("DPinvoice").Cells.Item(i + 1).Specific.Value;

                        if (invoiceDocEntry == docEntry)
                        {
                            row = i + 1;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentAlreadyAddedRow") + row);
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                            BubbleEvent = false;
                            break;
                        }

                    }
                    if (BubbleEvent == true)
                    {


                        oDocForm.Freeze(true);
                        oMatrix = ((SAPbouiCOM.Matrix)(oDocForm.Items.Item("DPinvoices").Specific));

                        oMatrix.AddRow(1, -1);
                        oMatrix.AutoResizeColumns();

                        decimal gTotal = Convert.ToDecimal(maxAmount);

                        SAPbouiCOM.EditText oEditText;

                        try
                        {
                            oEditText = oMatrix.Columns.Item("DPinvoice").Cells.Item(oMatrix.RowCount).Specific;
                            oEditText.Value = docEntry;
                        }
                        catch
                        { }
                        try
                        {
                            oEditText = oMatrix.Columns.Item("maxAmount").Cells.Item(oMatrix.RowCount).Specific;
                            oEditText.Value = maxAmount;
                        }
                        catch
                        { }
                        try
                        {
                            oEditText = oMatrix.Columns.Item("drgAmount").Cells.Item(oMatrix.RowCount).Specific;
                            oEditText.Value = maxAmount;
                        }
                        catch
                        { }
                        oMatrix.Columns.Item(0).Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                        setVisibleFormItems(oDocForm, out errorText);

                        oDocForm.Freeze(false);

                        oForm.Close();

                    }
                }

                if (pVal.ItemUID == "InvDocs" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                {
                    int row = pVal.Row;

                    SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    oForm.Freeze(true);
                    setInvDocsMatrixRowBackColor(oForm, row, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.ItemUID == "ChooseInvDocs" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvDocs").Specific));
                    SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
                    if (cellPos == null)
                    {
                        Program.uiApp.MessageBox(BDOSResources.getTranslate("ForPasswordSetChooseServiceUser"));
                        return;
                    }
                }

            }
            else if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems(oForm, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false)
                {
                    if (Program.FORM_LOAD_FOR_VISIBLE == true)
                    {
                        setSizeForm(oForm, out errorText);
                        oForm.Title = BDOSResources.getTranslate("TaxInvoiceReceived");
                        Program.FORM_LOAD_FOR_VISIBLE = false;
                    }
                    setVisibleFormItems(oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm(oForm, out errorText);
                    oForm.Freeze(false);
                }

                if ((pVal.ItemUID == "cardCodeE" || pVal.ItemUID == "confInfoE" || pVal.ItemUID == "wblMTR" || pVal.ItemUID == "DPinvoices" || pVal.ItemUID == "DPitems") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText, out BubbleEvent);
                }

                if (pVal.ItemUID == "wblMTR" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
                {
                    matrixColumnSetCfl(oForm, pVal, out errorText);
                }

                if (pVal.ItemUID == "wblMTR" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                {
                    matrixColumnSetCfl(oForm, pVal, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    itemPressed(oForm, pVal, out errorText);
                }

                if (pVal.ItemUID == "addMTRB" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    addMatrixRow(oForm, pVal.ItemUID, out errorText);
                }

                if (pVal.ItemUID == "addDPinv" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    chooseInvoiceDoc(oForm, "DownPaymentInvoice", out errorText);
                }

                if ((pVal.ItemUID == "delMTRB" || pVal.ItemUID == "delDPinv") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    delMatrixRow(oForm, pVal.ItemUID, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == true)
                {
                    //if (pVal.ItemUID == "wblMTR" & pVal.ColUID == "U_baseDocT")
                    //{
                    //    string downPaymnt = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_downPaymnt", 0).Trim();
                    //    string corrInv = oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("U_corrInv", 0).Trim();

                    //    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));
                    //    SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("U_baseDocT");

                    //    foreach (SAPbouiCOM.ValidValue oValidValue in oColumn.ValidValues)
                    //    {
                    //        if (oValidValue.Value == "0")
                    //            oMatrix.Columns.Item("U_baseDocT").ValidValues.Remove("0", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    //        if (oValidValue.Value == "1")
                    //            oMatrix.Columns.Item("U_baseDocT").ValidValues.Remove("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    //        if (oValidValue.Value == "2")
                    //            oMatrix.Columns.Item("U_baseDocT").ValidValues.Remove("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    //    }

                    //    if (downPaymnt == "Y")
                    //    {
                    //        oMatrix.Columns.Item("U_baseDocT").ValidValues.Add("2", BDOSResources.getTranslate("APDownPaymentInvoice"));
                    //    }
                    //    else if (corrInv == "Y")
                    //    {
                    //        oMatrix.Columns.Item("U_baseDocT").ValidValues.Add("1", BDOSResources.getTranslate("APCreditNote"));
                    //    }
                    //    else
                    //    {
                    //        oMatrix.Columns.Item("U_baseDocT").ValidValues.Add("0", BDOSResources.getTranslate("APInvoice"));
                    //        oMatrix.Columns.Item("U_baseDocT").ValidValues.Add("1", BDOSResources.getTranslate("APCreditNote"));
                    //    }
                    //}
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                {
                    comboSelect(oForm, pVal, out errorText);
                }

                //if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && pVal.BeforeAction == false
                //    && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE
                //    && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && pVal.EventType != SAPbouiCOM.BoEventTypes.et_GOT_FOCUS && pVal.EventType != SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                //    && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD
                //    && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.EventType != SAPbouiCOM.BoEventTypes.et_UDO_FORM_OPEN
                //    && pVal.EventType != SAPbouiCOM.BoEventTypes.et_VALIDATE)
                //{
                //    //setVisibleFormItems(oForm, out errorText);
                //}

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE & pVal.BeforeAction == false)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE == true)
                    {
                        setVisibleFormItems(oForm, out errorText);
                        formDataLoad(oForm, out errorText);
                        Program.FORM_LOAD_FOR_ACTIVATE = false;
                    }
                }

                if (pVal.ItemUID == "operationB" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == true)
                {
                    oForm.Freeze(true);
                    setValidValuesBtnCombo(oForm, out errorText);
                    oForm.Freeze(false);
                    //oForm.Update();
                }

                if (pVal.ItemUID == "postB" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        Program.uiApp.MessageBox(BDOSResources.getTranslate("ToCompleteOperationWriteDocument"));
                        return;
                    }
                    else if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationImpossibleInMode"));
                        return;
                    }

                    if (Program.uiApp.MessageBox(BDOSResources.getTranslate("DoYouWantToCreateJEForDownPayment") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "") == 1)
                    {
                        int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("DocEntry", 0));
                        postDocument(docEntry, out errorText);

                        if (!string.IsNullOrEmpty(errorText))
                            Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        FormsB1.SimulateRefresh();
                    }
                }
            }
        }

        private static void setInvDocsMatrixRowBackColor(SAPbouiCOM.Form oForm, int row, out string errorText)
        {
            errorText = null;

            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvDocs").Specific;

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
                GC.Collect();
            }
        }

        private static void setLinkStatus(SAPbobsCOM.GeneralData oGeneralData, SAPbobsCOM.GeneralDataCollection oChildren)
        {
            decimal amtBsDc = 0;
            decimal tAmtBsDc = 0;
            //oChildren = oGeneralData.Child("BDO_TXR1");

            foreach (SAPbobsCOM.GeneralData oChild in oChildren)
            {
                amtBsDc = amtBsDc + Convert.ToDecimal(oChild.GetProperty("U_amtBsDc")); //თანხა დღგ-ის ჩათვლით (საფუძველი დოკუმენტი)
                tAmtBsDc = tAmtBsDc + Convert.ToDecimal(oChild.GetProperty("U_tAmtBsDc")); //დღგ-ის თანხა (საფუძველი დოკუმენტი)
            }

            decimal amount = Convert.ToDecimal(oGeneralData.GetProperty("U_amount"));
            decimal amountTX = Convert.ToDecimal(oGeneralData.GetProperty("U_amountTX"));

            string linkStatus = getLinkStatus(amount, amountTX, amtBsDc, tAmtBsDc);
            oGeneralData.SetProperty("U_LinkStatus", linkStatus);
        }

        public static void postDocument(int docEntry, out string errorText, DateTime? taxDateFromTaxJournal = null)
        {
            errorText = null;

            SAPbobsCOM.CompanyService oCompanyService = Program.oCompany.GetCompanyService();
            SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXR_D");
            SAPbobsCOM.GeneralData oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

            SAPbobsCOM.GeneralDataParams oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            oGeneralParams.SetProperty("DocEntry", docEntry);
            oGeneralData = oGeneralService.GetByParams(oGeneralParams);

            bool corrInv = oGeneralData.GetProperty("U_corrInv") == "N" ? false : true;
            string invStatus = oGeneralData.GetProperty("U_status").Trim();
            //string docStatus = oGeneralData.GetProperty("Status");
            string jrnEnt = oGeneralData.GetProperty("U_JrnEnt");
            bool downPaymnt = oGeneralData.GetProperty("U_downPaymnt") == "Y";

            if (downPaymnt)
            {
                if (!string.IsNullOrEmpty(jrnEnt) && jrnEnt != "0")
                {
                    errorText = BDOSResources.getTranslate("JournalEntryAlreadyCreated") + "!";
                    return;
                }

                if (invStatus == "confirmed" || (invStatus == "corrected" && !corrInv) || invStatus == "correctionConfirmed")
                {
                    if (taxDateFromTaxJournal.HasValue)
                    {
                        oGeneralData.SetProperty("U_taxDate", taxDateFromTaxJournal.Value);
                    }

                    DateTime taxDate = oGeneralData.GetProperty("U_taxDate");
                    if (downPaymnt && taxDate == new DateTime(1899, 12, 30))
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("PleaseFillTaxDate"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return;
                    }

                    Program.JrnLinesGlobal = new DataTable();
                    DataTable JrnLinesDT = createAdditionalEntries(null, oGeneralData);
                    string DocNum = oGeneralData.GetProperty("DocNum").ToString();
                    DateTime DocDate = downPaymnt ? taxDate : oGeneralData.GetProperty("U_docDate");
                    string errorTextJrnEnt;
                    JrnEntry(docEntry.ToString(), DocNum, DocDate, JrnLinesDT, out errorTextJrnEnt);

                    if (errorTextJrnEnt != null)
                    {
                        errorText = BDOSResources.getTranslate("JournalEntryNotCreated") + "! " + BDOSResources.getTranslate("ReasonIs") + ": " + errorTextJrnEnt;
                    }
                    else
                    {
                        Program.JrnLinesGlobal = JrnLinesDT;
                        string Ref1 = docEntry.ToString();
                        string Ref2 = "UDO_F_BDO_TAXR_D";

                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = "SELECT " +
                                        "*  " +
                                        "FROM \"OJDT\"  " +
                                        "WHERE \"StornoToTr\" IS NULL " +
                                        "AND \"Ref1\" = '" + Ref1 + "' " +
                                        "AND \"Ref2\" = '" + Ref2 + "' ";
                        oRecordSet.DoQuery(query);

                        if (!oRecordSet.EoF)
                            oGeneralData.SetProperty("U_JrnEnt", Convert.ToString(oRecordSet.Fields.Item("TransId").Value, CultureInfo.InvariantCulture));
                        else
                            oGeneralData.SetProperty("U_JrnEnt", "");

                    }
                    oGeneralService.Update(oGeneralData);
                }
                else
                {
                    errorText = BDOSResources.getTranslate("CheckDocumentAndItsStatus") + "!";
                    return;
                }
            }
        }

        //--------------------------------------------RS.GE-------------------------------------------->
        public static void operationRS(TaxInvoice oTaxInvoice, string operation, int docEntry, int seqNum, DateTime DeclDate, DataRow taxDataRow, out string statusRS, out string errorText, bool fromTaxJournal = false)
        {
            errorText = null;
            statusRS = null;

            //CommonFunctions.StartTransaction;
            SAPbobsCOM.CompanyService oCompanyService = Program.oCompany.GetCompanyService();
            SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXR_D");
            SAPbobsCOM.GeneralData oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

            SAPbobsCOM.GeneralDataParams oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            oGeneralParams.SetProperty("DocEntry", docEntry);
            oGeneralData = oGeneralService.GetByParams(oGeneralParams);

            string status = oGeneralData.GetProperty("U_status");
            bool corrInv = oGeneralData.GetProperty("U_corrInv") == "N" ? false : true;
            string declNumber = oGeneralData.GetProperty("U_declNumber");

            if (operation == "updateStatus" || operation == "update" || operation == "create") //სტატუსების განახლება
            {
                if (operation == "update" && (status == "confirmed" || status == "correctionConfirmed"))
                    errorText = BDOSResources.getTranslate("UpdateConfirmedTaxInvoiceNotAllowed") + "!";
                else
                    get_invoice(oTaxInvoice, oGeneralService, oGeneralData, operation, taxDataRow, out errorText);
                if (errorText == null)
                {
                    oGeneralService.Update(oGeneralData);
                }
            }
            else if (operation == "confirmation") //დადასტურება
            {
                acsept_invoice_status(oTaxInvoice, oGeneralService, oGeneralData, out errorText);
                if (errorText == null)
                {
                    oGeneralService.Update(oGeneralData);
                }
            }
            else if (operation == "deny") //უარყოფა
            {
                ref_invoice_status(oTaxInvoice, oGeneralData, out errorText);
                if (errorText == null)
                {
                    oGeneralService.Update(oGeneralData);
                }
            }
            else if (operation == "addToTheDeclaration") //დეკლარაციაში დამატება
            {
                if (status == "confirmed" || status == "correctionConfirmed" || status == "corrected")
                {
                    add_inv_to_decl(oTaxInvoice, oGeneralData, seqNum, DeclDate, out errorText);
                }
                else
                {
                    errorText = BDOSResources.getTranslate("DocumentSStatusShouldBeTheOneFromTheList") + " : \"" + BDOSResources.getTranslate("Confirmed") + "\", \"" + BDOSResources.getTranslate("CorrectionConfirmed") + "\", \"" + BDOSResources.getTranslate("Corrected") + "\""; //დოკუმენტის სტატუსი უნდა იყოს ერთ-ერთი სიიდან
                }
                if (errorText == null)
                {
                    oGeneralService.Update(oGeneralData);
                }
            }
            else if (operation == "checkSync") //სინქრონიზაციის შემოწმება
            {
                if (checkSync(oTaxInvoice, oGeneralData, out statusRS, out errorText) == false)
                {
                    if (errorText == null)
                    {
                        errorText = BDOSResources.getTranslate("SynchronisationViolatedUpdateStatus");
                    }
                    return;
                }
            }

            //int docEntry = oGeneralData.GetProperty("DocEntry");
            //--------------------------------------------------------------------------------------------------------------------------------
            //string invoiceStatus = oGeneralData.GetProperty("U_status").Trim();
            //string jrnEnt = oGeneralData.GetProperty("U_JrnEnt");

            //string errorTextJrnEnt = null;

            //if ((operation == "updateStatus" || operation == "update" || operation == "create" || operation == "confirmation")
            //    && errorText == null && docEntry > 0 && (string.IsNullOrEmpty(jrnEnt) || jrnEnt == "0")
            //    && (invoiceStatus == "confirmed" || (invoiceStatus == "corrected" && !corrInv) || invoiceStatus == "correctionConfirmed"))
            //{
            //    string DocNum = oGeneralData.GetProperty("DocNum").ToString();

            //    Program.JrnLinesGlobal = new DataTable();
            //    DataTable reLines = null;

            //    if (!fromTaxJournal && oGeneralData.GetProperty("U_downPaymnt") == "Y")
            //    {
            //        DateTime taxDate = oGeneralData.GetProperty("U_taxDate");

            //        if (taxDate == new DateTime(1899, 12, 30))
            //        {
            //            createJournalEntryForTaxWithPostingDate = Program.uiApp.MessageBox(BDOSResources.getTranslate("DoYouWantCreateJournalEntryForTaxWithTaxInvoiceDate") + "? \n" + BDOSResources.getTranslate("IfNoPleaseFillTaxDateAndUpdateDocument"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
            //            if (createJournalEntryForTaxWithPostingDate == 2)
            //            {
            //                journalEntryForTaxWithPostingDateMsg = BDOSResources.getTranslate("ForCreateJournalEntryPleaseFillTaxDateAndUpdateDocument");
            //                return;
            //            }
            //        }
            //        else
            //            createJournalEntryForTaxWithPostingDate = 2;
            //    }

            //    if (fromTaxJournal && taxDatefromTaxJournal != null)
            //    {
            //        oGeneralData.SetProperty("U_taxDate", taxDatefromTaxJournal);
            //        createJournalEntryForTaxWithPostingDate = 2;
            //        //oGeneralService.Update(oGeneralData);
            //    }

            //    DataTable JrnLinesDT = createAdditionalEntries(null, oGeneralData);

            //    DateTime DocDate = createJournalEntryForTaxWithPostingDate == 2 ? oGeneralData.GetProperty("U_taxDate") : oGeneralData.GetProperty("U_docDate");
            //    JrnEntry(docEntry.ToString(), DocNum, DocDate, JrnLinesDT, reLines, out errorTextJrnEnt);
            //    if (errorTextJrnEnt != null)
            //    {
            //        errorText = BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + errorTextJrnEnt;
            //        oGeneralData.SetProperty("U_status", status); //ძველი სტატუსის დაბრუნება
            //    }
            //    else
            //    {
            //        Program.JrnLinesGlobal = JrnLinesDT;

            //        // გატარებები
            //        string Ref1 = docEntry.ToString();
            //        string Ref2 = "UDO_F_BDO_TAXR_D";

            //        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //        string query = "SELECT " +
            //                        "*  " +
            //                        "FROM \"OJDT\"  " +
            //                        "WHERE \"StornoToTr\" IS NULL " +
            //                        "AND \"Ref1\" = '" + Ref1 + "' " +
            //                        "AND \"Ref2\" = '" + Ref2 + "' ";
            //        oRecordSet.DoQuery(query);

            //        if (!oRecordSet.EoF)
            //        {
            //            oGeneralData.SetProperty("U_JrnEnt", Convert.ToString(oRecordSet.Fields.Item("TransId").Value, CultureInfo.InvariantCulture));
            //        }
            //        else
            //        {
            //            oGeneralData.SetProperty("U_JrnEnt", "");
            //        }
            //    }
            //    oGeneralService.Update(oGeneralData);
            //}
            //--------------------------------------------------------------------------------------------------------------------------------
            //if (Program.oCompany.InTransaction)
            //{
            //    //თუ დოკუმენტი გატარდა, მერე ვაკეთებს ბუღალტრულ გატარებას
            //    if (errorTextJrnEnt == null)
            //    {                    
            //        CommonFunctions.EndTransaction( SAPbobsCOM.BoWfTransOpt.wf_Commit);
            //    }
            //    else
            //    {
            //        CommonFunctions.EndTransaction( SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            //    }
            //}
            //else
            //{
            //    errorText = "ტრანზაქციის დასრულებს შეცდომა";
            //}
        }

        public static void linkToAPDocuments(int docEntry, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.CompanyService oCompanyService = Program.oCompany.GetCompanyService();
            SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXR_D");

            SAPbobsCOM.GeneralDataParams oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            oGeneralParams.SetProperty("DocEntry", docEntry);
            SAPbobsCOM.GeneralData oGeneralData = oGeneralService.GetByParams(oGeneralParams);

            SAPbobsCOM.GeneralDataCollection oChildren = oGeneralData.Child("BDO_TXR1");
            int oChildrenCount = oChildren.Count;
            bool searchByWbl = false;

            try
            {
                if (oGeneralData.GetProperty("U_downPaymnt") == "N")
                {
                    ///----------------------------------------------->დოკუმენტების მიბმა (შესყიდვა/დაბრუნება)<-----------------------------------------------
                    if (oChildrenCount != 0) //ზედნადების ნომრით, კონტრაგენტით, თარიღით
                    {

                        for (int i = 0; i < oChildrenCount; i++)
                        {
                            SAPbobsCOM.GeneralData oChildrenRow = oGeneralData.Child("BDO_TXR1").Item(i);

                            int baseDocEntry = oChildrenRow.GetProperty("U_baseDoc");
                            string overhead_no = oChildrenRow.GetProperty("U_wbNumber");
                            if (baseDocEntry <= 0 && string.IsNullOrEmpty(overhead_no) == false)
                            {
                                searchByWbl = true;
                                fillBaseDocs(oGeneralData, oChildren, oGeneralData.Child("BDO_TXR1").Item(i), overhead_no);
                            }
                        }
                    }
                }
                if (oChildrenCount == 0) //თანხით, კონტრაგენტით
                {
                    fillBaseDocs(oGeneralData, oChildren, null, null);
                }

                if (errorText == null && (oChildrenCount != oChildren.Count || searchByWbl))
                {
                    setLinkStatus(oGeneralData, oChildren);
                    oGeneralService.Update(oGeneralData);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.ReleaseComObject(oChildren);
                Marshal.ReleaseComObject(oGeneralParams);
                Marshal.ReleaseComObject(oGeneralData);
                Marshal.ReleaseComObject(oGeneralService);
            }
            ///----------------------------------------------->დოკუმენტების მიბმა (შესყიდვა/დაბრუნება)<-----------------------------------------------
        }

        private static void get_buyer_invoices(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralService oGeneralService, SAPbobsCOM.GeneralData oGeneralData, string operation, DataRow taxDataRow, int k_type, out string errorText)
        {
            errorText = null;

            try
            {
                string invoice_no = oGeneralData.GetProperty("U_number");
                DateTime opDate = oGeneralData.GetProperty("U_opDate");

                if (taxDataRow == null)
                {
                    DateTime firstDay = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                    DateTime lastDay = firstDay.AddMonths(1).AddDays(-1);

                    DateTime op_s_dt = DateTime.MinValue;
                    DateTime op_e_dt = lastDay;
                    DateTime s_dt = DateTime.MinValue;
                    DateTime e_dt = lastDay;

                    op_s_dt = DateTime.TryParse(opDate.ToString("yyyyMMdd") == "18991230" ? "" : opDate.ToString(), out op_s_dt) == false ? DateTime.MinValue : op_s_dt;
                    op_e_dt = DateTime.TryParse(opDate.ToString("yyyyMMdd") == "18991230" ? "" : opDate.ToString(), out op_s_dt) == false ? lastDay : op_e_dt;

                    DataTable taxDataTable = oTaxInvoice.get_buyer_invoices(s_dt, e_dt, op_s_dt, op_e_dt, invoice_no, "", "", "", out errorText);
                    if (taxDataTable.Rows.Count == 0)
                    {
                        //errorText = "ანგარიშ-ფაქტურა ამ პერიოდში არ მოიძებნა!"; //თუ რს-ზე იძებნება ა/ფ, მაშინ რადგან სტატუსი აქვს უარყოფილი რს-ზე მაგიტომ ვერ პოულობს ეს სერვისი (სავარაუდოდ).
                        return;
                    }
                    taxDataRow = taxDataTable.Rows[0];
                }

                object ID = taxDataRow["ID"]; // ანგარიშ-ფაქტურის უნიკალური ნომერი
                object SELLER_UN_ID = taxDataRow["SELLER_UN_ID"]; // გამყიდველის გადამხდელის უნიკალური ნომერი
                string SEQ_NUM_B = taxDataRow["SEQ_NUM_B"].ToString(); // მყიდველის დეკლარაციის ნომერი
                object STATUS = taxDataRow["STATUS"]; // ანგარიშ-ფაქტურის სტატუსი
                int WAS_REF = Convert.ToInt32(string.IsNullOrEmpty(taxDataRow["WAS_REF"].ToString()) ? 0 : taxDataRow["WAS_REF"]); // უარყოფილი მეორე მხარის მიერ 0 - არა 1 - კი
                object F_SERIES = taxDataRow["F_SERIES"]; // ანგარიშ-ფაქტურის სერია
                object F_NUMBER = taxDataRow["F_NUMBER"]; // ანგარიშ-ფაქტურის ნომერი
                DateTime REG_DT = new DateTime();
                REG_DT = DateTime.TryParse(taxDataRow["REG_DT"].ToString(), out REG_DT) == false ? new DateTime() : REG_DT; // რეგისტრაციის თარიღი
                DateTime OPERATION_DT = new DateTime();
                OPERATION_DT = DateTime.TryParse(taxDataRow["OPERATION_DT"].ToString(), out OPERATION_DT) == false ? new DateTime() : OPERATION_DT; // ოპერაციის განხორციელების თარიღი
                object S_USER_ID = taxDataRow["S_USER_ID"]; // სერვისის მომხმარებლის უნიკალური ნომერი
                object B_S_USER_ID = taxDataRow["B_S_USER_ID"]; // მყიდველის სერვისის მომხმარებლის უნიკალური ნომერი
                object DOC_MOS_NOM_B = taxDataRow["DOC_MOS_NOM_B"]; // ??? 
                object SA_IDENT_NO = taxDataRow["SA_IDENT_NO"]; // მყიდველის საიდენტიფიკაციო ნომერი
                object ORG_NAME = taxDataRow["ORG_NAME"]; // მყიდველის დასახელება 
                object NOTES = taxDataRow["NOTES"]; // მყიდველის მაღაზიის ნომერი
                decimal TANXA = Convert.ToDecimal(string.IsNullOrEmpty(taxDataRow["TANXA"].ToString()) ? 0 : taxDataRow["TANXA"], CultureInfo.InvariantCulture); // თანხა  დღგ-ის ჩათვლით
                decimal VAT = Convert.ToDecimal(string.IsNullOrEmpty(taxDataRow["VAT"].ToString()) ? 0 : taxDataRow["VAT"], CultureInfo.InvariantCulture); // დღგ-ის თანხა
                string K_ID = taxDataRow["K_ID"].ToString(); // კორექტირების ანგარიშ-ფაქტურის ID
                DateTime AGREE_DATE = new DateTime();
                AGREE_DATE = DateTime.TryParse(taxDataRow["AGREE_DATE"].ToString(), out AGREE_DATE) == false ? new DateTime() : AGREE_DATE; // დადასტურების თარიღი
                object AGREE_S_USER_ID = taxDataRow["AGREE_S_USER_ID"]; // დამდასტურებელი
                DateTime REF_DATE = new DateTime();
                REF_DATE = DateTime.TryParse(taxDataRow["REF_DATE"].ToString(), out REF_DATE) == false ? new DateTime() : REF_DATE; // უარყოფის თარიღი
                object REF_S_USER_ID = taxDataRow["REF_S_USER_ID"]; // უარმყოფელი

                string downPaymnt = "N";
                string corrInv = "N";

                if (F_SERIES.ToString().Contains("ავ"))
                    downPaymnt = "Y";
                else if (F_SERIES.ToString().Contains("აკ"))
                {
                    downPaymnt = "Y";
                    corrInv = "Y";
                }
                else if (F_SERIES.ToString().Contains("ეკ"))
                    corrInv = "Y";
                else if (F_SERIES.ToString().Contains("ეა"))
                {
                }

                if (STATUS != null)
                {
                    STATUS = (WAS_REF == 1) ? "0" : STATUS; //თუ არის უარყოფილი ა/ფ                   
                    STATUS = getStatusValueByStatusNumber(STATUS.ToString());

                    oGeneralData.SetProperty("U_status", STATUS);
                    oGeneralData.SetProperty("U_declNumber", SEQ_NUM_B);
                    oGeneralData.SetProperty("U_invID", ID.ToString());
                    oGeneralData.SetProperty("U_number", F_NUMBER.ToString());
                    oGeneralData.SetProperty("U_series", F_SERIES.ToString());
                    oGeneralData.SetProperty("U_recvDate", REG_DT);
                    oGeneralData.SetProperty("U_opDate", OPERATION_DT);
                    oGeneralData.SetProperty("U_confDate", AGREE_DATE);
                    oGeneralData.SetProperty("U_downPaymnt", downPaymnt);
                    oGeneralData.SetProperty("U_corrInv", corrInv);

                    DateTime emptyDate = new DateTime(1, 1, 1);
                    if (AGREE_DATE == emptyDate)
                    {
                        DateTime createDate = oGeneralData.GetProperty("CreateDate");
                        oGeneralData.SetProperty("U_docDate", createDate);
                    }
                    else
                    {
                        oGeneralData.SetProperty("U_docDate", AGREE_DATE);
                    }

                    if (operation == "update" || operation == "create") //განახლება
                    {
                        decimal amtCor = 0; //თანხა დღგ-ის ჩათვლით (კორექტირებული ა/ფ)
                        decimal amtTXCor = 0; //დღგ-ის თანხა (კორექტირებული ა/ფ)
                        int corrDocEntry;
                        List<string> wbNumbersCor = new List<string>();
                        List<string> wbNumbers = new List<string>();

                        SAPbobsCOM.GeneralDataCollection oChildren = null;

                        if (corrInv == "Y") //თუ არის კორექტირების ა/ფ (რასაც აკორექტირებს იმ ა/ფ-ის ID)
                        {
                            oGeneralData.SetProperty("U_corrDocID", K_ID);
                            oGeneralData.SetProperty("U_corrType", k_type == -1 ? "-1" : k_type.ToString());

                            getInfoDoc(K_ID, null, out corrDocEntry, out amtCor, out amtTXCor, out wbNumbersCor, out errorText);

                            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string query = @"SELECT *
                            FROM ""@BDO_TXR5""
                            WHERE ""@BDO_TXR5"".""DocEntry"" =  '" + corrDocEntry + "'";
                            oRecordSet.DoQuery(query);

                            if (oRecordSet.RecordCount > 0)
                            {
                                oChildren = oGeneralData.Child("BDO_TXR5");

                                while (!oRecordSet.EoF)
                                {
                                    double vatAmount = oRecordSet.Fields.Item("U_drg_amount").Value;
                                    string tax_invoice = oRecordSet.Fields.Item("U_tax_invoice").Value;
                                    oRecordSet.MoveNext();

                                    SAPbobsCOM.GeneralData oChild = null;
                                    oChild = oChildren.Add();
                                    oChild.SetProperty("U_drg_amount", vatAmount);
                                    oChild.SetProperty("U_tax_invoice", tax_invoice);
                                }
                            }

                            oGeneralData.SetProperty("U_corrDoc", corrDocEntry);
                            oGeneralData.SetProperty("U_corrDTxt", corrDocEntry.ToString());

                            //კორექტირებული ფაქტურის მონაცემები ---->
                            oGeneralData.SetProperty("U_amtCor", Convert.ToDouble(amtCor)); //თანხა დღგ-ის ჩათვლით
                            oGeneralData.SetProperty("U_amtTXCor", Convert.ToDouble(amtTXCor)); //დღგ-ის თანხა

                            //მონაცემები კორექტირების შემდეგ ---->
                            oGeneralData.SetProperty("U_amtACor", Convert.ToDouble(TANXA)); //თანხა დღგ-ის ჩათვლით
                            oGeneralData.SetProperty("U_amtTXACr", Convert.ToDouble(VAT)); //დღგ-ის თანხა
                        }

                        decimal amount = TANXA - amtCor;
                        decimal amountTX = VAT - amtTXCor;
                        oGeneralData.SetProperty("U_amount", Convert.ToDouble(amount));
                        oGeneralData.SetProperty("U_amountTX", Convert.ToDouble(amountTX));

                        string tin = SA_IDENT_NO.ToString();
                        string cardName;
                        string cardCode;

                        cardCode = BusinessPartners.GetCardCodeByTin(tin, "S", out cardName);
                        if (cardCode == null)
                        {
                            errorText = BDOSResources.getTranslate("BPNotFound") + " " + BDOSResources.getTranslate("BPTin") + " : " + tin;
                            return;
                        }
                        oGeneralData.SetProperty("U_cardCode", cardCode);
                        oGeneralData.SetProperty("U_cardCodeN", cardName);
                        oGeneralData.SetProperty("U_cardCodeT", tin);

                        oChildren = null;

                        int id = 0; //ზედნადების ჩანაწერის უნიკალური ID
                        int inv_id = 0; //ანგარიშ-ფაქტურის უნიკალური ნომერი
                        string overhead_no = null; //ზედნადების ნომერი
                        DateTime overhead_dt = new DateTime(); //ზედნადების თარიღი
                        string overhead_dt_str = null; //ზედნადების თარიღი (სტრიქონი)
                        bool searchByWbl = false;

                        if (downPaymnt == "N")
                        {
                            //ზედნადების ცხრილური ნაწილი 
                            DataTable invoiceTableLines = oTaxInvoice.get_ntos_invoices_inv_nos(Convert.ToInt32(ID), out errorText);

                            oChildren = oGeneralData.Child("BDO_TXR2"); //მხოლოდ ზედნადების ნომრების ცხრილი
                            while (oChildren.Count > 0)
                            {
                                oChildren.Remove(0);
                            }
                            for (int i = 0; i < invoiceTableLines.Rows.Count; i++)
                            {
                                SAPbobsCOM.GeneralData oChildGeneralData = oChildren.Add();
                                DataRow invoiceRow = invoiceTableLines.Rows[i];
                                oChildGeneralData.SetProperty("U_wbNumber", invoiceRow["overhead_no"].ToString());
                            }

                            ///----------------------------------------------->დოკუმენტების მიბმა (შესყიდვა/დაბრუნება)<-----------------------------------------------
                            if (invoiceTableLines.Rows.Count != 0) //ზედნადების ნომრით, კონტრაგენტით, თარიღით
                            {
                                oChildren = oGeneralData.Child("BDO_TXR1");

                                for (int i = 0; i < invoiceTableLines.Rows.Count; i++)
                                {
                                    DataRow invoiceRow = invoiceTableLines.Rows[i];
                                    id = Convert.ToInt32(invoiceRow["id"]); //ზედნადების ჩანაწერის უნიკალური ID
                                    inv_id = Convert.ToInt32(invoiceRow["inv_id"]); //ანგარიშ-ფაქტურის უნიკალური ნომერი
                                    overhead_no = invoiceRow["overhead_no"].ToString(); //ზედნადების ნომერი
                                    overhead_dt = Convert.ToDateTime(invoiceRow["overhead_dt"]); //ზედნადების თარიღი
                                    overhead_dt_str = invoiceRow["overhead_dt_str"].ToString(); //ზედნადების თარიღი (სტრიქონი)                                  

                                    if (wbNumbersCor.Contains(overhead_no) == true)
                                    {
                                        continue;
                                    }
                                    if (searchByWbl == false)
                                    {
                                        while (oChildren.Count > 0)
                                        {
                                            oChildren.Remove(0);
                                        }
                                        oGeneralService.Update(oGeneralData);
                                    }
                                    searchByWbl = true;
                                    fillBaseDocs(oGeneralData, oChildren, null, overhead_no);
                                }
                            }
                        }
                        if (searchByWbl == false && operation == "create") //თანხით, კონტრაგენტით
                        {
                            oChildren = oGeneralData.Child("BDO_TXR1");
                            fillBaseDocs(oGeneralData, oChildren, null, null);
                        }
                        ///----------------------------------------------->დოკუმენტების მიბმა (შესყიდვა/დაბრუნება)<-----------------------------------------------

                        if (downPaymnt != "N")
                        {
                            //ავანსის ფაქტურისთვის ცხრილის შევსება

                            int inv_ID = Convert.ToInt32(ID);

                            oChildren = oGeneralData.Child("BDO_TXR3"); //აითემების ცხრილი
                            while (oChildren.Count > 0)
                            {
                                oChildren.Remove(0);
                            }

                            DataTable taxDataTable = oTaxInvoice.get_invoice_desc(inv_ID, out errorText);

                            for (int i = 0; i < taxDataTable.Rows.Count; i++)
                            {

                                SAPbobsCOM.GeneralData oChildGeneralData = oChildren.Add();

                                DataRow taxDeclRow = taxDataTable.Rows[i];

                                //try
                                //{
                                //    string taxID = taxDeclRow["id"].ToString(); //ID
                                //    oChildGeneralData.SetProperty("id", taxID);
                                //}
                                //catch { }
                                try
                                {
                                    decimal g_number = Convert.ToDecimal(taxDeclRow["g_number"], CultureInfo.InvariantCulture); //რაოდენობა
                                    oChildGeneralData.SetProperty("U_g_number", Convert.ToDouble(g_number));
                                }
                                catch { }
                                try
                                {
                                    decimal full_amount = Convert.ToDecimal(taxDeclRow["full_amount"], CultureInfo.InvariantCulture); //თანხა დღგ-ის და აქციზის ჩათვლით
                                    oChildGeneralData.SetProperty("U_full_amount", Convert.ToDouble(full_amount));
                                }
                                catch { }
                                try
                                {
                                    decimal drg_amount = Convert.ToDecimal(taxDeclRow["drg_amount"], CultureInfo.InvariantCulture); //დღგ
                                    oChildGeneralData.SetProperty("U_drg_amount", Convert.ToDouble(drg_amount));
                                }
                                catch { }
                                try
                                {
                                    string goods = taxDeclRow["goods"].ToString();
                                    oChildGeneralData.SetProperty("U_goods", goods);
                                }
                                catch { }
                                try
                                {
                                    string g_unit = taxDeclRow["g_unit"].ToString();
                                    oChildGeneralData.SetProperty("U_g_unit", g_unit);
                                }
                                catch { }
                                try
                                {
                                    string RSVatCode = taxDeclRow["vat_type"].ToString();

                                    oChildGeneralData.SetProperty("U_vat_type", BDO_WaybillsJournalReceived.DetectVATByRSCode(RSVatCode, out errorText));
                                }
                                catch { }
                            }
                            //ავანსის ფაქტურისთვის ცხრილის შევსება
                        }

                        //decimal amtBsDc = 0;
                        //decimal tAmtBsDc = 0;
                        //oChildren = oGeneralData.Child("BDO_TXR1");

                        //foreach (SAPbobsCOM.GeneralData oChild in oChildren)
                        //{
                        //    amtBsDc = amtBsDc + Convert.ToDecimal(oChild.GetProperty("U_amtBsDc")); //თანხა დღგ-ის ჩათვლით (საფუძველი დოკუმენტი)
                        //    tAmtBsDc = tAmtBsDc + Convert.ToDecimal(oChild.GetProperty("U_tAmtBsDc")); //დღგ-ის თანხა (საფუძველი დოკუმენტი)
                        //}

                        //string linkStatus = getLinkStatus(amount, amountTX, amtBsDc, tAmtBsDc);
                        //oGeneralData.SetProperty("U_LinkStatus", linkStatus);
                        oChildren = oGeneralData.Child("BDO_TXR1");
                        setLinkStatus(oGeneralData, oChildren);
                    }
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

        private static void fillBaseDocs(SAPbobsCOM.GeneralData oGeneralData, SAPbobsCOM.GeneralDataCollection oChildren, SAPbobsCOM.GeneralData oChildGeneralData, string overhead_no)
        {
            string errorText = null;
            DataTable baseDocs;

            baseDocs = getListBaseDoc(oGeneralData, overhead_no, null, 0, out errorText);

            if (string.IsNullOrEmpty(overhead_no) == false) //ვეძებთ ზედნადების ნომრით
            {
                int baseDocsCount = baseDocs.Rows.Count;

                if (baseDocsCount > 0)
                {
                    for (int i = 0; i < baseDocsCount; i++)
                    {
                        DataRow dataRow = baseDocs.Rows[i];
                        if (oChildGeneralData == null)
                        {
                            oChildGeneralData = oChildren.Add();
                            oChildGeneralData.SetProperty("U_wbNumber", overhead_no);
                        }

                        int baseDocEntry = Convert.ToInt32(dataRow["DocEntry"]);
                        oChildGeneralData.SetProperty("U_baseDocT", Convert.ToString(dataRow["BaseDocType"]));
                        oChildGeneralData.SetProperty("U_baseDoc", baseDocEntry);
                        oChildGeneralData.SetProperty("U_baseDTxt", baseDocEntry.ToString());
                        oChildGeneralData.SetProperty("U_amtBsDc", Convert.ToDouble(dataRow["GTotal"]));
                        oChildGeneralData.SetProperty("U_tAmtBsDc", Convert.ToDouble(dataRow["LineVat"]));
                    }
                }
                else if (oChildGeneralData == null)
                {
                    oChildGeneralData = oChildren.Add();
                    oChildGeneralData.SetProperty("U_wbNumber", overhead_no);
                }
            }
            else //ვეძებთ თანხის მიხედვით
            {
                if (baseDocs.Rows.Count == 1)
                {
                    DataRow dataRow = baseDocs.Rows[0];

                    oChildGeneralData = oChildren.Add();
                    int baseDocEntry = Convert.ToInt32(dataRow["DocEntry"]);
                    oChildGeneralData.SetProperty("U_baseDocT", Convert.ToString(dataRow["BaseDocType"]));
                    oChildGeneralData.SetProperty("U_baseDoc", baseDocEntry);
                    oChildGeneralData.SetProperty("U_baseDTxt", baseDocEntry.ToString());
                    oChildGeneralData.SetProperty("U_amtBsDc", Convert.ToDouble(dataRow["GTotal"]));
                    oChildGeneralData.SetProperty("U_tAmtBsDc", Convert.ToDouble(dataRow["LineVat"]));
                }
            }
        }

        /// <summary>რეკვიზიტების განახლება</summary>
        private static void get_invoice(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralService oGeneralService, SAPbobsCOM.GeneralData oGeneralData, string operation, DataRow taxDataRow, out string errorText)
        {
            errorText = null;

            try
            {
                string invID = oGeneralData.GetProperty("U_invID");
                if (taxDataRow != null)
                {
                    invID = taxDataRow["ID"].ToString();
                }

                if (String.IsNullOrEmpty(invID) == true)
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceIDNotFilled");
                    return;
                }

                int inv_ID = Convert.ToInt32(invID);

                bool corrInv = oGeneralData.GetProperty("U_corrInv") == "N" ? false : true;
                Dictionary<string, object> responseDictionary = oTaxInvoice.get_invoice(inv_ID, out errorText); //(- არ აბრუნებს დადასტურების თარიღს, არ აბრუნებს უარყოფილია თუ არა (წაშლილი ა/ფ ჩანს))
                if (errorText != null)
                {
                    return;
                }

                bool result = Convert.ToBoolean(responseDictionary["result"]);
                if (result == false)
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceNotFoundOnSite") + errorText;
                    return;
                }
                else
                {
                    DateTime reg_dt = Convert.ToDateTime(responseDictionary["reg_dt"]);
                    int f_number = Convert.ToInt32(responseDictionary["f_number"]);
                    string f_series = responseDictionary["f_series"] == null ? "" : responseDictionary["f_series"].ToString();
                    int statusRS = Convert.ToInt32(responseDictionary["status"]);
                    string seq_num_b = responseDictionary["seq_num_b"] == null ? "" : responseDictionary["seq_num_b"].ToString();
                    string seq_num_s = responseDictionary["seq_num_s"] == null ? "" : responseDictionary["seq_num_s"].ToString();
                    DateTime operation_dt = Convert.ToDateTime(responseDictionary["operation_dt"]);
                    int seller_un_id = Convert.ToInt32(responseDictionary["seller_un_id"]);
                    int buyer_un_id = Convert.ToInt32(responseDictionary["buyer_un_id"]);
                    string overhead_no = responseDictionary["overhead_no"] == null ? "" : responseDictionary["overhead_no"].ToString();
                    int k_id = Convert.ToInt32(responseDictionary["k_id"]);
                    int r_un_id = Convert.ToInt32(responseDictionary["r_un_id"]);
                    int k_type = Convert.ToInt32(responseDictionary["k_type"]);
                    int b_s_user_id = Convert.ToInt32(responseDictionary["b_s_user_id"]);
                    int dec_status = Convert.ToInt32(responseDictionary["dec_status"]);

                    string status = getStatusValueByStatusNumber(statusRS.ToString());

                    oGeneralData.SetProperty("U_invID", invID);
                    oGeneralData.SetProperty("U_status", status);
                    oGeneralData.SetProperty("U_opDate", operation_dt);
                    oGeneralData.SetProperty("U_recvDate", reg_dt);
                    oGeneralData.SetProperty("U_number", f_number == -1 ? "" : f_number.ToString());
                    oGeneralData.SetProperty("U_series", f_series);
                    oGeneralData.SetProperty("U_declNumber", seq_num_b);
                    //oGeneralData.SetProperty("U_corrType", k_type == -1 ? "-1" : k_type.ToString());

                    string invoice_no = oGeneralData.GetProperty("U_number");
                    DateTime opDate = oGeneralData.GetProperty("U_opDate");

                    if (String.IsNullOrEmpty(invoice_no) == false && opDate != new DateTime())
                    {
                        get_buyer_invoices(oTaxInvoice, oGeneralService, oGeneralData, operation, taxDataRow, k_type, out errorText);
                    }
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

        /// <summary>სინქრონიზაციის შემოწმება (სტატუსის მიხედვით)</summary>
        public static bool checkSync(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, out string statusRS, out string errorText)
        {
            errorText = null;
            statusRS = null;

            try
            {
                string invID = oGeneralData.GetProperty("U_invID");
                if (String.IsNullOrEmpty(invID) == true)
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceIDNotFilled");
                    return false;
                }

                int inv_ID = Convert.ToInt32(invID);

                bool corrInv = oGeneralData.GetProperty("U_corrInv") == "N" ? false : true;
                Dictionary<string, object> responseDictionary = oTaxInvoice.get_invoice(inv_ID, out errorText); //(- არ აბრუნებს დადასტურების თარიღს, არ აბრუნებს უარყოფილია თუ არა (წაშლილი ა/ფ ჩანს))
                if (errorText != null)
                {
                    return false;
                }

                bool result = Convert.ToBoolean(responseDictionary["result"]);
                if (result == false)
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceNotFoundOnSite") + errorText;
                    return false;
                }
                else
                {
                    statusRS = responseDictionary["status"].ToString();
                    string status = getStatusValueByStatusNumber(statusRS);
                    string statusDoc = oGeneralData.GetProperty("U_status");
                    if (status == statusDoc)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return false;
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>უარყოფა</summary>
        private static void ref_invoice_status(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, out string errorText)
        {
            errorText = null;

            try
            {
                string statusRS;
                if (checkSync(oTaxInvoice, oGeneralData, out statusRS, out errorText) == false)
                {
                    if (errorText == null)
                    {
                        errorText = BDOSResources.getTranslate("SynchronisationViolatedUpdateStatus");
                    }
                    return;
                }

                string invID = oGeneralData.GetProperty("U_invID");
                string comment = oGeneralData.GetProperty("U_comment");

                if (String.IsNullOrEmpty(invID) == true)
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceIDNotFilled");
                    return;
                }

                int inv_ID = Convert.ToInt32(invID);
                bool response = oTaxInvoice.ref_invoice_status(inv_ID, comment, out errorText);
                if (response == true)
                {
                    oGeneralData.SetProperty("U_status", "denied"); //უარყოფილი
                }
                else
                {
                    errorText = BDOSResources.getTranslate("Operation") + " \"" + BDOSResources.getTranslate("RSDeny") + "\" " + BDOSResources.getTranslate("DoneWithErrors");
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

        /// <summary>დადასტურება</summary>
        private static void acsept_invoice_status(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralService oGeneralService, SAPbobsCOM.GeneralData oGeneralData, out string errorText)
        {
            errorText = null;

            try
            {
                string statusRS;
                if (checkSync(oTaxInvoice, oGeneralData, out statusRS, out errorText) == false)
                {
                    if (errorText == null)
                    {
                        errorText = BDOSResources.getTranslate("SynchronisationViolatedUpdateStatus");
                    }
                    return;
                }
                if (statusRS == "2")
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceAlreadyConfirmed");
                    return;
                }

                int newStatus = 0;

                if (statusRS == "6") //გაუქმების პროცესში
                {
                    newStatus = 7; //გაუქმებული
                }
                else if (statusRS == "1") //მიღებული (დასადასტურებელი)
                {
                    newStatus = 2; //დადასტურებული
                }
                else if (statusRS == "5") //კორექტირება მიღებული (დასადასტურებელი)
                {
                    newStatus = 8; //კორექტირება დადასტურებული
                }

                string invID = oGeneralData.GetProperty("U_invID");

                if (String.IsNullOrEmpty(invID) == true)
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceIDNotFilled");
                    return;
                }

                int inv_ID = Convert.ToInt32(invID);
                bool response = false;

                if (newStatus == 7)  //გაუქმებული
                {
                    response = oTaxInvoice.acsept_invoice_status(inv_ID, newStatus, out errorText);
                }
                else if (newStatus == 2) //დადასტურებული
                {
                    decimal amount = Convert.ToDecimal(oGeneralData.GetProperty("U_amount")); //თანხა დღგ-ის ჩათვლით
                    decimal amountTX = Convert.ToDecimal(oGeneralData.GetProperty("U_amountTX")); //დღგ-ის თანხა
                    decimal full_amount;
                    decimal drg_amount;

                    get_invoice_desc(oTaxInvoice, oGeneralData, out full_amount, out drg_amount, out errorText);

                    if (errorText != null)
                    {
                        return;
                    }

                    if (amount != full_amount || amountTX != drg_amount)
                    {
                        errorText = BDOSResources.getTranslate("TaxInvoiceAmountsNotEqualSiteNotUpdated");
                        return;
                    }

                    response = oTaxInvoice.acsept_invoice_status(inv_ID, newStatus, out errorText);
                }
                else if (newStatus == 8) //კორექტირება დადასტურებული
                {
                    decimal amtACor = Convert.ToDecimal(oGeneralData.GetProperty("U_amtACor")); //თანხა დღგ-ის ჩათვლით (კორექტირების შემდეგ)
                    decimal amtTXACr = Convert.ToDecimal(oGeneralData.GetProperty("U_amtTXACr")); //დღგ-ის თანხა (კორექტირების შემდეგ)
                    decimal full_amount;
                    decimal drg_amount;

                    get_invoice_desc(oTaxInvoice, oGeneralData, out full_amount, out drg_amount, out errorText);

                    if (errorText != null)
                    {
                        return;
                    }

                    if (amtACor != full_amount || amtTXACr != drg_amount)
                    {
                        errorText = BDOSResources.getTranslate("TaxInvoiceAmountsNotEqualSiteNotUpdated");
                        return;
                    }

                    response = oTaxInvoice.acsept_invoice_status(inv_ID, newStatus, out errorText);
                }

                if (response == true)
                {
                    oGeneralData.SetProperty("U_status", getStatusValueByStatusNumber(newStatus.ToString()));

                    if (newStatus == 2 || newStatus == 8) //დადასტურებული || კორექტირება დადასტურებული
                    {
                        get_invoice(oTaxInvoice, oGeneralService, oGeneralData, "updateStatus", null, out errorText);

                        //if (newStatus == 2)
                        //{
                        //დამდასტურებელი პირის შევსება
                        string userName = "";
                        int userID = 0;
                        Users.getUserByCode(Program.oCompany.UserName, out userName, out userID, out errorText);
                        if (errorText == null)
                        {
                            oGeneralData.SetProperty("U_confInfo", userID.ToString());
                            oGeneralData.SetProperty("U_confInfN", userName);
                        }
                        //}
                    }
                }
                else
                {
                    errorText = BDOSResources.getTranslate("Operation") + " \"" + BDOSResources.getTranslate("RSConfirm") + "\" " + BDOSResources.getTranslate("DoneWithErrors");
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

        /// <summary>"თანხა დღგ-ის და აქციზის ჩათვლით" და "დღგ" - ს საიტიდან მიღება</summary>
        private static void get_invoice_desc(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, out decimal full_amount, out decimal drg_amount, out string errorText)
        {
            errorText = null;
            decimal g_number = 0;
            full_amount = 0;
            drg_amount = 0;
            decimal aqcizi_amount = 0;

            try
            {
                string invID = oGeneralData.GetProperty("U_invID");

                if (String.IsNullOrEmpty(invID) == true)
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceIDNotFilled");
                    return;
                }

                int inv_ID = Convert.ToInt32(invID);

                DataTable taxDataTable = oTaxInvoice.get_invoice_desc(inv_ID, out errorText);

                for (int i = 0; i < taxDataTable.Rows.Count; i++)
                {
                    DataRow taxDeclRow = taxDataTable.Rows[i];
                    try
                    {
                        g_number = g_number + Convert.ToDecimal(taxDeclRow["g_number"], CultureInfo.InvariantCulture); //რაოდენობა
                    }
                    catch { }
                    try
                    {
                        full_amount = full_amount + Convert.ToDecimal(taxDeclRow["full_amount"], CultureInfo.InvariantCulture); //თანხა დღგ-ის და აქციზის ჩათვლით
                    }
                    catch { }
                    try
                    {
                        drg_amount = drg_amount + Convert.ToDecimal(taxDeclRow["drg_amount"], CultureInfo.InvariantCulture); //დღგ
                    }
                    catch { }
                    try
                    {
                        aqcizi_amount = aqcizi_amount + Convert.ToDecimal(taxDeclRow["aqcizi_amount"], CultureInfo.InvariantCulture); //აქციზის თანხა
                    }
                    catch { }
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

        /// <summary>დეკლარაციაში დამატება</summary>
        private static void add_inv_to_decl(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, int seqNum, DateTime declDate, out string errorText)
        {
            errorText = null;

            try
            {
                string statusRS;
                int docEntry = oGeneralData.GetProperty("DocEntry");

                if (checkSync(oTaxInvoice, oGeneralData, out statusRS, out errorText) == false)
                {
                    if (errorText == null)
                    {
                        errorText = BDOSResources.getTranslate("SynchronisationViolatedUpdateStatus");
                    }
                    return;
                }

                string invID = oGeneralData.GetProperty("U_invID");

                if (String.IsNullOrEmpty(invID) == true)
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceIDNotFilled");
                    return;
                }

                if (String.IsNullOrEmpty(oGeneralData.GetProperty("U_declNumber")) == false)
                {
                    errorText = BDOSResources.getTranslate("DeclarationNumberShouldBeEmpty");
                    return;
                }

                int inv_ID = Convert.ToInt32(invID);

                if (seqNum == -1) //დეკლარაციის ნომრების მიღება -->
                {
                    declDate = oGeneralData.GetProperty("U_declDate");

                    if (declDate.ToString("yyyyMMdd") == "18991230")
                    {
                        errorText = BDOSResources.getTranslate("DeclDateNotFilled");
                        return;
                    }

                    string period = new DateTime(declDate.Year, declDate.Month, 1).ToString("yyyyMM");

                    DataTable taxDeclTable = oTaxInvoice.get_seq_nums(period, out errorText);
                    if (errorText != null)
                    {
                        errorText = BDOSResources.getTranslate("CantReceiveDeclarationDate") + " " + declDate;
                        return;
                    }

                    for (int i = 0; i < taxDeclTable.Rows.Count; i++)
                    {
                        DataRow TaxDeclRow = taxDeclTable.Rows[i];
                        seqNum = Convert.ToInt32(TaxDeclRow.ItemArray[0]);
                    }

                    if (seqNum == 0)
                    {
                        errorText = BDOSResources.getTranslate("CantReceiveDeclarationDate") + " " + declDate;
                        return;
                    }
                } //<--

                //დეკლარაციაში დამატება
                bool response = oTaxInvoice.add_inv_to_decl(seqNum, inv_ID, out errorText);

                if (response == true)
                {
                    oGeneralData.SetProperty("U_declNumber", seqNum.ToString());
                    oGeneralData.SetProperty("U_declDate", declDate);
                }
                else
                {
                    errorText = BDOSResources.getTranslate("Operation") + " \"" + BDOSResources.getTranslate("RSAddDeclaration") + "\" " + BDOSResources.getTranslate("DoneWithErrors");
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
        //<--------------------------------------------RS.GE--------------------------------------------
    }
}