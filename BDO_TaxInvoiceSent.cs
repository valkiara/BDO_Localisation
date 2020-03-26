using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using SAPbobsCOM;

namespace BDO_Localisation_AddOn
{
    static partial class BDO_TaxInvoiceSent
    {
        const int clientHeight = 650;
        const int clientWidth = 800;

        public static void createDocumentUDO(out string errorText)
        {
            errorText = null;
            string tableName = "BDO_TAXS";
            string description = "Tax Invoice Sent";

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
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Customer Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ბიზნესპარტნიორი (სახელი)
            fieldskeysMap.Add("Name", "cardCodeN");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Customer Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ბიზნესპარტნიორი (გსნ)
            fieldskeysMap.Add("Name", "cardCodeT");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Customer TIN");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //გაგზავნის თარიღი
            fieldskeysMap.Add("Name", "sentDate");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Sent Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ID
            fieldskeysMap.Add("Name", "invID");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Invoice ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ნომერი
            fieldskeysMap.Add("Name", "number");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Invoice Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //სერია
            fieldskeysMap.Add("Name", "series");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Invoice Series");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>(); //სტატუსი
            listValidValuesDict.Add("empty", "");
            listValidValuesDict.Add("created", "Created"); //შექმნილი
            listValidValuesDict.Add("shipped", "Shipped"); //გადაგზავნილი
            listValidValuesDict.Add("confirmed", "Confirmed"); //დადასტურებული
            listValidValuesDict.Add("removed", "Removed"); //წაშლილი
            listValidValuesDict.Add("incompleteShipped", "Created Incompletely"); //არასრულად შექმნილი
            listValidValuesDict.Add("paper", "Paper"); //ქაღალდის
            listValidValuesDict.Add("disturbedSynchronization", "Disturbed Synchronization"); //დარღვეულია სინქრონიზაცია
            listValidValuesDict.Add("denied", "Denied"); //უარყოფილი
            listValidValuesDict.Add("cancellationProcess", "Cancellation Process"); //გაუქმების პროცესში
            listValidValuesDict.Add("canceled", "Canceled"); //გაუქმებული
            listValidValuesDict.Add("attachedToTheDeclaration", "Attached To The Declaration"); //დეკლარაციაზე მიბმული
            listValidValuesDict.Add("correctionCreated", "Correction Created"); //შექმნილი კორექტირებული
            listValidValuesDict.Add("correctionShipped", "Correction Shipped"); //გადაგზავნილი კორექტირებული
            listValidValuesDict.Add("correctionConfirmed", "Correction Confirmed"); //დადასტურებული კორექტირებული
            listValidValuesDict.Add("primary", "Primary"); //პირველადი
            listValidValuesDict.Add("corrected", "Corrected"); //კორექტირებული

            fieldskeysMap = new Dictionary<string, object>(); //სტატუსი
            fieldskeysMap.Add("Name", "status");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Invoice Status");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დადასტურების თარიღი
            fieldskeysMap.Add("Name", "confDate");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Confirmation Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დეკლარაციის თვე
            fieldskeysMap.Add("Name", "declDate");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Declaration Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დეკლარაციის ნომერი
            fieldskeysMap.Add("Name", "declNumber");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Declaration Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ოპერაციის თვე
            fieldskeysMap.Add("Name", "opDate");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Operation Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amount"); //თანხა დღგ-ის ჩათვლით
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amountTX"); //დღგ-ის თანხა
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Vat Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtOutTX"); //თანხა დღგ-ის გარეშე 
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Amount Without Vat");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //მაკორექტირებელი ანგარიშ–ფაქტურისთვის
            fieldskeysMap.Add("Name", "corrInv");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "For Correction");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            listValidValues = new List<string>(); //კორექტირების მიზეზები
            listValidValues.Add("");
            listValidValues.Add("Canceled Tax Operation"); //1 //გაუქმებულია დასაბეგრი ოპერაცია
            listValidValues.Add("Changed Tax Operation Type"); //2 //შეცვლილია დასაბეგრი ოპერაციის სახე
            listValidValues.Add("Changed Agreement Amount Prices Decrease"); //3 //ფასების შემცირების ან სხვა მიზეზით შეცვლილია ოპერაციაზე ადრე შეთანხმებული კომპენსაციის თანხა
            listValidValues.Add("Item Service Returned To Seller"); //4 საქონელი (მომსახურება) სრულად ან ნაწილობრივ უბრუნდება გამყიდველს

            fieldskeysMap = new Dictionary<string, object>(); //კორექტირების მიზეზები
            fieldskeysMap.Add("Name", "corrType");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Correction Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ელექტრონული სახით
            fieldskeysMap.Add("Name", "elctrnic");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Electronic");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "Y");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ჩათვლის თვე
            fieldskeysMap.Add("Name", "vatRDate");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Vat Receive Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //კომენტარი
            fieldskeysMap.Add("Name", "comment");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Comment");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "corrDoc"); //კორექტირების დოკუმენტი
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Correction Document");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "corrDTxt"); //კორექტირების დოკუმენტი (TXT)
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Correction Document");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "corrDocID"); //კორექტირების დოკუმენტის ID
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Correction Document ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ავანსის ანგარიშ–ფაქტურა
            fieldskeysMap.Add("Name", " downPaymnt");
            fieldskeysMap.Add("TableName", "BDO_TAXS");
            fieldskeysMap.Add("Description", "Down Payment");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            tableName = "BDO_TXS1";
            description = "Tax Invoice Sent Child1";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            if (result != 0)
            {
                return;
            }

            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("ARInvoice", "ARInvoice");
            listValidValuesDict.Add("ARCreditNote", "ARCreditNote");
            listValidValuesDict.Add("ARCorrectionInvoice", "ARCorrectionInvoice");
            listValidValuesDict.Add("ARDownPaymentRequest", "ARDownPaymentRequest");
            listValidValuesDict.Add("ARDownPaymentInvoice", "ARDownPaymentInvoice");
            listValidValuesDict.Add("ARDownPaymentVAT", "ARDownPaymentVAT");

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "baseDocT"); //საფუძველი დოკუმენტის ტიპი
            fieldskeysMap.Add("TableName", "BDO_TXS1");
            fieldskeysMap.Add("Description", "Base Document Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            bool result2 = UDO.addNewValidValuesUserFieldsMD("@BDO_TXS1", "baseDocT", "ARDownPaymentRequest", "ARDownPaymentRequest", out errorText);
            result2 = UDO.addNewValidValuesUserFieldsMD("@BDO_TXS1", "baseDocT", "ARDownPaymentVAT", "ARDownPaymentVAT", out errorText);
            result2 = UDO.addNewValidValuesUserFieldsMD("@BDO_TXS1", "baseDocT", "ARCorrectionInvoice", "ARCorrectionInvoice", out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "baseDoc"); //საფუძველი დოკუმენტი
            fieldskeysMap.Add("TableName", "BDO_TXS1");
            fieldskeysMap.Add("Description", "Base Document");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "baseDTxt"); //საფუძველი დოკუმენტი (TXT)
            fieldskeysMap.Add("TableName", "BDO_TXS1");
            fieldskeysMap.Add("Description", "Base Document");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtBsDc"); //საფუძველი დოკუმენტის თანხა
            fieldskeysMap.Add("TableName", "BDO_TXS1");
            fieldskeysMap.Add("Description", "Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "tAmtBsDc"); //საფუძველი დოკუმენტის დღგ-ის თანხა
            fieldskeysMap.Add("TableName", "BDO_TXS1");
            fieldskeysMap.Add("Description", "Vat Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "wbNumber"); //ზედნადების ნომერი
            fieldskeysMap.Add("TableName", "BDO_TXS1");
            fieldskeysMap.Add("Description", "Waybill Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void registerUDO(out string errorText)
        {
            errorText = null;
            string code = "UDO_F_BDO_TAXS_D"; //20 characters (must include at least one alphabetical character).
            Dictionary<string, object> formProperties;

            formProperties = new Dictionary<string, object>();
            formProperties.Add("Name", "Tax Invoice Sent"); //100 characters
            formProperties.Add("TableName", "BDO_TAXS");
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
            fieldskeysMap.Add("ColumnDescription", "Customer Code"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_cardCodeN");
            fieldskeysMap.Add("ColumnDescription", "Customer Name"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_cardCodeT");
            fieldskeysMap.Add("ColumnDescription", "Customer TIN"); //30 characters
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
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "CreateDate");
            fieldskeysMap.Add("ColumnDescription", "Create Date"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "UpdateDate");
            fieldskeysMap.Add("ColumnDescription", "Update Date"); //30 characters
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
            fieldskeysMap.Add("TableName", "BDO_TXS1");
            fieldskeysMap.Add("ObjectName", "BDO_TXS1"); //30 characters
            listChildTables.Add(fieldskeysMap);

            formProperties.Add("ChildTables", listChildTables);

            UDO.registerUDO(code, formProperties, out errorText);

            GC.Collect();
        }

        public static void addMenus()
        {
            try
            {
                SAPbouiCOM.MenuItem fatherMenuItem = Program.uiApp.Menus.Item("2048");
                // Add a pop-up menu item
                SAPbouiCOM.MenuCreationParams oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "UDO_F_BDO_TAXS_D";
                oCreationPackage.String = BDOSResources.getTranslate("TaxInvoiceSent");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                SAPbouiCOM.MenuItem menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems;
            string itemName = "";

            int left_s = 6;
            int left_e = 127;
            int height = 15;
            int top = 6;
            int width_s = 121;
            int width_e = 148;
            int fontSize = 10;

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "cardCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CustomerCode"));
            formItems.Add("LinkTo", "cardCodeE");
            formItems.Add("FontSize", fontSize);

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
            oCon.CondVal = "C"; //მყიდველი
            oCFL.SetConditions(oCons);

            formItems = new Dictionary<string, object>();
            itemName = "cardCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
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
            formItems.Add("TableName", "@BDO_TAXS");
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
            formItems.Add("Caption", BDOSResources.getTranslate("CustomerTIN"));
            formItems.Add("LinkTo", "cardCodeTE");
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "cardCodeTE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
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

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "sentDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Sent"));
            formItems.Add("LinkTo", "sentDateE");
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "sentDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
            formItems.Add("Alias", "U_sentDate");
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
            formItems.Add("TableName", "@BDO_TAXS");
            formItems.Add("Alias", "U_corrInv");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e * 1.26);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Caption", BDOSResources.getTranslate("CorrectionTaxInvoice"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "elctrnicCH"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
            formItems.Add("Alias", "U_elctrnic");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Left", left_s + width_e * 1.26);
            formItems.Add("Width", width_e * 2 / 3);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Caption", BDOSResources.getTranslate("Electronic"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("FontSize", fontSize);

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
            formItems.Add("TableName", "@BDO_TAXS");
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
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //კორექტირების დოკუმენტის ტიპის მიხედვით CFL- ის დამატება ---->           
            multiSelection = false;
            objectType = "UDO_F_BDO_TAXS_D"; //Tax Invoice Sent
            string uniqueID_CorrDocCFL = "CorrDoc_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_CorrDocCFL);
            //<----

            left_s = 300;
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
            formItems.Add("TableName", "@BDO_TAXS");
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

            formItems = new Dictionary<string, object>();
            itemName = "corrDocE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
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

            formItems = new Dictionary<string, object>();
            itemName = "corrDocE1"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
            formItems.Add("Alias", "U_corrDoc");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 2 + 20);
            formItems.Add("Width", width_e / 2 - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            //formItems.Add("Enabled", false);
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
            formItems.Add("TableName", "@BDO_TAXS");
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
            formItems.Add("Enabled", true);
            formItems.Add("Visible", false);
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = 6;

            formItems = new Dictionary<string, object>();
            itemName = "No.S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s / 3);
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
            formItems.Add("TableName", "@BDO_TAXS");
            formItems.Add("Alias", "Series");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_s + width_s / 3);
            formItems.Add("Width", width_s * 2 / 3);
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
            formItems.Add("TableName", "@BDO_TAXS");
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
            formItems.Add("TableName", "@BDO_TAXS");
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
            formItems.Add("TableName", "@BDO_TAXS");
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

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "UpdateDatS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("UpdateDate"));
            formItems.Add("LinkTo", "UpdateDatE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "UpdateDatE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
            formItems.Add("Alias", "UpdateDate");
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

            left_s = 6;
            left_e = 127;

            //ელექტრონული ანგარიშ-ფაქტურა ---->

            top = 70 + 40;

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
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "seriesE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
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

            formItems = new Dictionary<string, object>();
            itemName = "numberE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
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

            top = top + height + 1;

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
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "invIDE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
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
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("empty", "");
            listValidValuesDict.Add("created", BDOSResources.getTranslate("Created"));
            listValidValuesDict.Add("shipped", BDOSResources.getTranslate("Sent"));
            listValidValuesDict.Add("confirmed", BDOSResources.getTranslate("Confirmed"));
            listValidValuesDict.Add("removed", BDOSResources.getTranslate("deleted"));
            listValidValuesDict.Add("incompleteShipped", BDOSResources.getTranslate("CreatedIncompletely"));
            listValidValuesDict.Add("paper", BDOSResources.getTranslate("Paper"));
            listValidValuesDict.Add("disturbedSynchronization", BDOSResources.getTranslate("SynchronizationViolated"));
            listValidValuesDict.Add("denied", BDOSResources.getTranslate("Denied"));
            listValidValuesDict.Add("cancellationProcess", BDOSResources.getTranslate("CancellationProcess"));
            listValidValuesDict.Add("canceled", BDOSResources.getTranslate("Canceled"));
            listValidValuesDict.Add("attachedToTheDeclaration", BDOSResources.getTranslate("RSAddDeclaration"));
            listValidValuesDict.Add("correctionCreated", BDOSResources.getTranslate("CreatedCorrected"));
            listValidValuesDict.Add("correctionShipped", BDOSResources.getTranslate("SentCorrected"));
            listValidValuesDict.Add("correctionConfirmed", BDOSResources.getTranslate("ConfirmedCorrected"));
            listValidValuesDict.Add("primary", BDOSResources.getTranslate("Primary"));
            listValidValuesDict.Add("corrected", BDOSResources.getTranslate("Corrected"));

            formItems = new Dictionary<string, object>();
            itemName = "statusCB"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
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
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "confDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
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
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "declDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
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
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "declNmberE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
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

            left_s = 300;
            left_e = left_s + 121;
            top = 70 + 40;
            top = top + 25;

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
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "opDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
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
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "amountE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
            formItems.Add("Alias", "U_amount");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            //formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

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
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "amountTXE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
            formItems.Add("Alias", "U_amountTX");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            //formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "amtOutTXS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("AmountWithoutVAT"));
            formItems.Add("LinkTo", "amtOutTXE");
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "amtOutTXE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
            formItems.Add("Alias", "U_amtOutTX");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            //formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

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
            formItems.Add("FontSize", fontSize);

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
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            left_s = 6;
            left_e = 127;

            top = top + 2 * height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "vatRDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("VATReceiveDate"));
            formItems.Add("LinkTo", "vatRDateE");
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "vatRDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXS");
            formItems.Add("Alias", "U_vatRDate");
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

            top = top + 2 * height + 1;

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

            int wblMTRWidth = oForm.ClientWidth;

            formItems = new Dictionary<string, object>();
            itemName = "wblMTR"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            formItems.Add("Left", left_s);
            formItems.Add("Top", top);
            formItems.Add("UID", itemName);

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
            objectType = "13"; //SAPbouiCOM.BoLinkedObject.lf_Invoice
            string uniqueID_lf_InvoiceCFL = "Invoice_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_InvoiceCFL);

            objectType = "14"; //SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo 
            string uniqueID_lf_InvoiceCreditMemoCFL = "InvoiceCreditMemo_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_InvoiceCreditMemoCFL);

            objectType = "165";
            string uniqueID_lf_InvoiceCorrectionInvoiceCFL = "InvoiceCorrectionInvoice_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_InvoiceCorrectionInvoiceCFL);

            objectType = "203"; //A/R Down Payment Invoice
            string uniqueID_lf_DownPaymentInvoiceCFL = "DownPaymentInvoice_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_DownPaymentInvoiceCFL);

            objectType = "UDO_F_BDO_ARDPV_D"; //A/R Down Payment Invoice
            string uniqueID_lf_ARDownPaymentVAT_CFL = "ARDownPaymentVAT_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_ARDownPaymentVAT_CFL);

            oColumn = oColumns.Add("LineId", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Editable = false;

            wblMTRWidth = wblMTRWidth - 20 - 1;

            SAPbobsCOM.UserTable oUserTable = Program.oCompany.UserTables.Item("BDO_TXS1");
            SAPbobsCOM.ValidValues oValidValues = oUserTable.UserFields.Fields.Item("U_baseDocT").ValidValues;

            oColumn = oColumns.Add("U_baseDocT", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocumentType");
            oColumn.Editable = true;

            foreach (SAPbobsCOM.ValidValue keyValue in oValidValues)
            {
                oColumn.ValidValues.Add(keyValue.Value, BDOSResources.getTranslate(keyValue.Value));
            }

            oColumn = oColumns.Add("U_baseDoc", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Document");
            oColumn.Editable = true;
            oColumn.Visible = false;

            oColumn = oColumns.Add("U_baseDTxt", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Document");
            oColumn.Editable = true;

            oColumn = oColumns.Add("U_amtBsDc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("AmountWithVat");
            oColumn.Editable = false;

            oColumn = oColumns.Add("U_tAmtBsDc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatAmount");
            oColumn.Editable = false;

            oColumn = oColumns.Add("U_wbNumber", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("WaybillNumber");
            oColumn.Editable = false;

            SAPbouiCOM.DBDataSource oDBDataSource;
            oDBDataSource = oForm.DataSources.DBDataSources.Add("@BDO_TXS1");

            oColumn = oColumns.Item("LineId");
            oColumn.DataBind.SetBound(true, "@BDO_TXS1", "LineId");

            oColumn = oColumns.Item("U_baseDocT");
            oColumn.DataBind.SetBound(true, "@BDO_TXS1", "U_baseDocT");
            oColumn.DisplayDesc = true;
            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

            oColumn = oColumns.Item("U_baseDoc");
            oColumn.DataBind.SetBound(true, "@BDO_TXS1", "U_baseDoc");

            oColumn = oColumns.Item("U_baseDTxt");
            oColumn.DataBind.SetBound(true, "@BDO_TXS1", "U_baseDTxt");

            oColumn = oColumns.Item("U_amtBsDc");
            oColumn.DataBind.SetBound(true, "@BDO_TXS1", "U_amtBsDc");

            oColumn = oColumns.Item("U_tAmtBsDc");
            oColumn.DataBind.SetBound(true, "@BDO_TXS1", "U_tAmtBsDc");

            oColumn = oColumns.Item("U_wbNumber");
            oColumn.DataBind.SetBound(true, "@BDO_TXS1", "U_wbNumber");

            //oMatrix.Clear();
            //oDBDataSource.Query();
            //oMatrix.LoadFromDataSource();
            //oMatrix.AutoResizeColumns();

            //სარდაფი
            left_s = 6;
            left_e = 127;
            top = top + oForm.Items.Item("wblMTR").Height + 40;

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
            formItems.Add("TableName", "@BDO_TAXS");
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
            formItems.Add("TableName", "@BDO_TAXS");
            formItems.Add("Alias", "Remark");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
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
            formItems.Add("TableName", "@BDO_TAXS");
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

            left_s = 300;
            left_e = left_s + 121;

            top = top + height + 1;

            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("RSUpdateStatus"));
            listValidValuesDict.Add("save", BDOSResources.getTranslate("Save"));
            listValidValuesDict.Add("send", BDOSResources.getTranslate("Send"));
            listValidValuesDict.Add("remove", BDOSResources.getTranslate("Delete"));
            listValidValuesDict.Add("cancel", BDOSResources.getTranslate("Cancel"));
            listValidValuesDict.Add("addToTheDeclaration", BDOSResources.getTranslate("RSAddDeclaration"));

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

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                int height = 15;
                int top = 6;
                top = top + height + 1;

                SAPbouiCOM.Item oItem = oForm.Items.Item("cardCodeS");
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

                oItem = oForm.Items.Item("sentDateS");
                oItem.Top = top;
                oItem = oForm.Items.Item("sentDateE");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("corrInvCH");
                oItem.Top = top;
                oItem = oForm.Items.Item("elctrnicCH");
                oItem.Top = top;
                oItem = oForm.Items.Item("corrDocIDS");
                oItem.Top = top;
                oItem = oForm.Items.Item("corrDocIDE");
                oItem.Top = top;
                oItem = oForm.Items.Item("corrDocE");
                oItem.Top = top;
                oItem = oForm.Items.Item("corrDocLB");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("downPmntCH");
                oItem.Top = top;
                top = top + height + 1;

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

                oItem = oForm.Items.Item("UpdateDatS");
                oItem.Top = top;
                oItem = oForm.Items.Item("UpdateDatE");
                oItem.Top = top;
                top = top + height + 1;

                //ელექტრონული ანგარიშ-ფაქტურა ---->

                top = 70 + 40;

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

                oItem = oForm.Items.Item("invIDS");
                oItem.Top = top;
                oItem = oForm.Items.Item("invIDE");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("statusS1");
                oItem.Top = top;
                oItem = oForm.Items.Item("statusCB");
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

                top = 70 + 40;
                top = top + 25;

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
                top = top + height + 1;

                oItem = oForm.Items.Item("amtOutTXS");
                oItem.Top = top;
                oItem = oForm.Items.Item("amtOutTXE");
                oItem.Top = top;
                top = top + height + 1;

                top = oForm.Items.Item("declNmberS").Top;
                oItem = oForm.Items.Item("declStatS");
                oItem.Top = top;
                oItem = oForm.Items.Item("declStatS1");
                oItem.Top = top;
                top = top + 2 * height + 1;

                oItem = oForm.Items.Item("vatRDateS");
                oItem.Top = top;
                oItem = oForm.Items.Item("vatRDateE");
                oItem.Top = top;
                top = top + 2 * height + 1;

                oItem = oForm.Items.Item("addMTRB");
                oItem.Top = top;
                oItem = oForm.Items.Item("delMTRB");
                oItem.Top = top;
                top = top + height + 1;

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("wblMTR").Specific;
                int mtrWidth = oForm.ClientWidth;
                oForm.Items.Item("wblMTR").Width = mtrWidth;
                oForm.Items.Item("wblMTR").Top = top;
                oForm.Items.Item("wblMTR").Height = oForm.ClientHeight / 3;
                FormsB1.resetWidthMatrixColumns(oForm, "wblMTR", "LineId", mtrWidth);

                if (oForm.ClientHeight <= clientHeight)
                    top += oForm.Items.Item("wblMTR").Height + 10;
                else
                    top = oForm.ClientHeight - (5 * height + 10);

                oForm.Items.Item("CreatorS").Top = top;
                oForm.Items.Item("CreatorE").Top = top;

                top += height + 1;

                oForm.Items.Item("RemarksS").Top = top;
                oForm.Items.Item("RemarksE").Top = top;

                top += height + 1;

                oForm.Items.Item("CommentS").Top = top;
                oForm.Items.Item("CommentE").Top = top;

                top += 2 * height + 1;

                oForm.Items.Item("1").Top = top;
                oForm.Items.Item("2").Top = top;

                oItem = oForm.Items.Item("operationB");
                oItem.Left = oForm.ClientWidth - 6 - oItem.Width;
                oItem.Top = top;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void setSizeForm(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                oForm.ClientHeight = clientHeight;
                oForm.ClientWidth = clientWidth;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Item oItem = null;
            oForm.Freeze(true);

            try
            {
                string docEntry = oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("DocEntry", 0).Trim();
                bool docEntryIsEmpty = string.IsNullOrEmpty(docEntry);
                oForm.Items.Item("corrInvCH").Enabled = (docEntryIsEmpty);
                oForm.Items.Item("downPmntCH").Enabled = (docEntryIsEmpty);

                string corrInv = oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("U_corrInv", 0).Trim();
                oItem = oForm.Items.Item("corrDocIDS");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("corrDocIDE");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("corrDocE");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("corrDocLB");
                oItem.Visible = corrInv == "N" ? false : true;
                oItem = oForm.Items.Item("corrTypeCB");
                oItem.Visible = corrInv == "N" ? false : true;

                setValidValuesBtnCombo(oForm, out errorText);

                string elctrnic = oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("U_elctrnic", 0).Trim();
                if (elctrnic == "N")
                {
                    oItem = oForm.Items.Item("seriesE"); //სერია
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("numberE"); //ნომერი
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("invIDE"); //ID
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("statusCB"); //ID
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("confDateE"); //დადასტურების თარიღი
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("declNmberE"); //დეკლარაციის ნომერი
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("opDateE"); //ოპერაციის თვე
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("amountE"); //თანხა დღგ-ის ჩათვლით
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("amountTXE"); //დღგ-ის თანხა
                    oItem.Enabled = true;
                    oItem = oForm.Items.Item("amtOutTXE"); //თანხა დღგ-ის გარეშე
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
                    oItem = oForm.Items.Item("statusCB"); //ID
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("confDateE"); //დადასტურების თარიღი
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("declNmberE"); //დეკლარაციის ნომერი
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("corrDocIDE"); //კორექტირების დოკუმენტის ID
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("corrDocE"); //კორექტირების დოკუმენტი
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("opDateE"); //ოპერაციის თვე
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("amountE"); //თანხა დღგ-ის ჩათვლით
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("amountTXE"); //დღგ-ის თანხა
                    oItem.Enabled = false;
                    oItem = oForm.Items.Item("amtOutTXE"); //თანხა დღგ-ის გარეშე
                    oItem.Enabled = false;
                }

                string invoiceStatus = oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("U_status", 0).Trim();
                oItem = oForm.Items.Item("operationB");

                if (invoiceStatus == "paper") // ქაღალდის               
                    oItem.Visible = false;
                else if (invoiceStatus != "")
                    oItem.Visible = true;
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
                SAPbouiCOM.ButtonCombo oButtonCombo = ((SAPbouiCOM.ButtonCombo)(oForm.Items.Item("operationB").Specific));

                if (pVal.BeforeAction)
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

                        if (pVal.ItemChanged)
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

                            if (selectedOperation == "updateStatus") //სტატუსების განახლება
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("DocEntry", 0));
                                string errorTextWb;
                                string errorTextGoods;
                                operationRS(oTaxInvoice, selectedOperation, docEntry, -1, new DateTime(), out errorText, out errorTextWb, out errorTextGoods);
                                if (errorText != null)
                                {
                                    Program.uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("RSUpdateStatus") + " " + BDOSResources.getTranslate("DoneSuccessfully"));
                                }
                            }
                            else if (selectedOperation == "save") //შენახვა
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("DocEntry", 0));
                                string errorTextWb;
                                string errorTextGoods;
                                operationRS(oTaxInvoice, selectedOperation, docEntry, -1, new DateTime(), out errorText, out errorTextWb, out errorTextGoods);
                                if (errorText != null)
                                {
                                    Program.uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("Save") + " " + BDOSResources.getTranslate("DoneSuccessfully"));
                                }
                            }
                            else if (selectedOperation == "send") //გადაგზავნა
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("DocEntry", 0));
                                string errorTextWb;
                                string errorTextGoods;
                                operationRS(oTaxInvoice, selectedOperation, docEntry, -1, new DateTime(), out errorText, out errorTextWb, out errorTextGoods);
                                if (errorText != null)
                                {
                                    Program.uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("Send") + " " + BDOSResources.getTranslate("DoneSuccessfully"));
                                }
                            }
                            else if (selectedOperation == "remove") //წაშლა
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("DocEntry", 0));
                                string errorTextWb;
                                string errorTextGoods;
                                operationRS(oTaxInvoice, selectedOperation, docEntry, -1, new DateTime(), out errorText, out errorTextWb, out errorTextGoods);
                                if (errorText != null)
                                {
                                    Program.uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("Delete") + " " + BDOSResources.getTranslate("DoneSuccessfully"));
                                }
                            }
                            else if (selectedOperation == "cancel") //გაუქმება
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("DocEntry", 0));
                                string errorTextWb;
                                string errorTextGoods;
                                operationRS(oTaxInvoice, selectedOperation, docEntry, -1, new DateTime(), out errorText, out errorTextWb, out errorTextGoods);
                                if (errorText != null)
                                {
                                    Program.uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("Cancel") + " " + BDOSResources.getTranslate("DoneSuccessfully"));
                                }
                            }
                            else if (selectedOperation == "addToTheDeclaration") //დეკლარაციაში დამატება
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("DocEntry", 0));
                                string errorTextWb;
                                string errorTextGoods;
                                operationRS(oTaxInvoice, selectedOperation, docEntry, -1, new DateTime(), out errorText, out errorTextWb, out errorTextGoods);
                                if (errorText != null)
                                {
                                    Program.uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("RSAddDeclaration") + " " + BDOSResources.getTranslate("DoneSuccessfully"));
                                }
                            }
                            if (selectedOperation != null)
                            {
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
                string declNumber = oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("U_declNumber", 0).Trim();
                SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("declStatS1").Specific;
                if (declNumber == "")
                {
                    oStaticText.Caption = BDOSResources.getTranslate("NotLinked");
                }
                else
                {
                    oStaticText.Caption = BDOSResources.getTranslate("Linked");
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

        public static void setValidValuesBtnCombo(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            try
            {
                Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
                string invoiceStatus = oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("U_status", 0).Trim();
                SAPbouiCOM.Item oItem = oForm.Items.Item("operationB");

                if (invoiceStatus == "paper" || invoiceStatus == "canceled" || invoiceStatus == "removed") // ქაღალდის || გაუქმებული || წაშლილი
                {
                    oItem.Visible = false;
                }
                else if (invoiceStatus == "empty" || invoiceStatus == "") //ცარიელი || ""
                {
                    oItem.Visible = true;
                    listValidValuesDict.Add("save", BDOSResources.getTranslate("Save"));
                    listValidValuesDict.Add("send", BDOSResources.getTranslate("Send"));
                }
                else if (invoiceStatus == "created" || invoiceStatus == "correctionCreated" || invoiceStatus == "incompleteShipped") //შექმნილი || შექმნილი კორექტირებული || არასრულად შექმნილი
                {
                    oItem.Visible = true;
                    listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("RSUpdateStatus"));
                    listValidValuesDict.Add("save", BDOSResources.getTranslate("Save"));
                    listValidValuesDict.Add("send", BDOSResources.getTranslate("Send"));
                    listValidValuesDict.Add("remove", BDOSResources.getTranslate("Delete"));
                }
                else if (invoiceStatus == "shipped" || invoiceStatus == "correctionShipped") //გადაგზავნილი || გადაგზავნილი კორექტირებული
                {
                    oItem.Visible = true;
                    listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("RSUpdateStatus"));
                    listValidValuesDict.Add("send", BDOSResources.getTranslate("Send"));
                    listValidValuesDict.Add("remove", BDOSResources.getTranslate("Delete"));
                }
                else if (invoiceStatus == "confirmed" || invoiceStatus == "correctionConfirmed") //დადასტურებული || დადასტურებული კორექტირებული
                {
                    oItem.Visible = true;
                    listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("RSUpdateStatus"));
                    if (oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("U_declNumber", 0).Trim() == "")
                    {
                        listValidValuesDict.Add("cancel", BDOSResources.getTranslate("Cancel"));
                        listValidValuesDict.Add("addToTheDeclaration", BDOSResources.getTranslate("RSAddDeclaration"));
                    }
                }
                else if (invoiceStatus == "denied") //უარყოფილი 
                {
                    oItem.Visible = true;
                    listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("RSUpdateStatus"));
                    listValidValuesDict.Add("save", BDOSResources.getTranslate("Save"));
                    listValidValuesDict.Add("send", BDOSResources.getTranslate("Send"));
                    listValidValuesDict.Add("remove", BDOSResources.getTranslate("Delete"));
                }
                else if ((invoiceStatus == "primary" || invoiceStatus == "corrected") && oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("U_declNumber", 0).Trim() == "") // (პირველადი || კორექტირებული) && დეკლარაციის ნომერი არ არის შევსებული
                {
                    oItem.Visible = true;
                    listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("RSUpdateStatus"));
                    listValidValuesDict.Add("addToTheDeclaration", BDOSResources.getTranslate("RSAddDeclaration"));
                }
                else
                {
                    oItem.Visible = true;
                    listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("RSUpdateStatus"));
                }
                if (oItem.Visible)
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
                    if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED & pVal.BeforeAction) || (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS & pVal.BeforeAction == false))
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));

                        SAPbouiCOM.ComboBox oComboBox = oMatrix.Columns.Item("U_baseDocT").Cells.Item(pVal.Row).Specific;
                        SAPbouiCOM.Column oColumn;

                        if (oComboBox.Value == "ARInvoice") //რეალიზაცია
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            oColumn.ChooseFromListUID = "Invoice_CFL";
                            oColumn.ChooseFromListAlias = "DocEntry";
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "13"; //SAPbouiCOM.BoLinkedObject.lf_Invoice
                        }
                        else if (oComboBox.Value == "ARCreditNote") //კორექტირება
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            oColumn.ChooseFromListUID = "InvoiceCreditMemo_CFL";
                            oColumn.ChooseFromListAlias = "DocEntry";
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "14"; //SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo
                        }
                        else if (oComboBox.Value == "ARCorrectionInvoice") //AR Correction Invoice
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            oColumn.ChooseFromListUID = "InvoiceCorrectionInvoice_CFL";
                            oColumn.ChooseFromListAlias = "DocEntry";
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "165";
                        }
                        else if (oComboBox.Value == "ARDownPaymentRequest") //ავანსი
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            oColumn.ChooseFromListUID = "DownPaymentInvoice_CFL";
                            oColumn.ChooseFromListAlias = "DocEntry";
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "203"; //A/R Down Payment Invoice
                        }
                        else if (oComboBox.Value == "ARDownPaymentVAT") //ავანსი
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            oColumn.ChooseFromListUID = "ARDownPaymentVAT_CFL";
                            oColumn.ChooseFromListAlias = "DocEntry";
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "UDO_F_BDO_ARDPV_D"; //A/R Down Payment Invoice
                        }

                    }
                }
                else
                {

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
                if (pVal.ItemUID == "elctrnicCH")
                {
                    string elctrnic = oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("U_elctrnic", 0).Trim();
                    if (elctrnic == "N")
                    {
                        oForm.DataSources.DBDataSources.Item("@BDO_TAXS").SetValue("U_status", 0, "paper");
                    }
                    else
                    {
                        oForm.DataSources.DBDataSources.Item("@BDO_TAXS").SetValue("U_status", 0, "empty");
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
                    //FormsB1.resetWidthMatrixColumns(oForm, "wblMTR", "LineId", wblMTRWidth);
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

        public static void selectMatrixRow(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (pVal.ColUID == "LineId")
                {
                    if (pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));
                        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                        oMatrix.SelectRow(pVal.Row, true, true);
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

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction, int row, out string errorText)
        {
            errorText = null;

            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (beforeAction)
                {
                    if (sCFL_ID == "Invoice_CFL" || sCFL_ID == "InvoiceCreditMemo_CFL" || sCFL_ID == "InvoiceCorrectionInvoice_CFL" || sCFL_ID == "DownPaymentInvoice_CFL" || sCFL_ID == "ARDownPaymentVAT_CFL")
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
                            string cardCode = oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("U_cardCode", 0).Trim();

                            int docEntry = 0;

                            try
                            {
                                docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("docEntry", 0));
                            }
                            catch
                            {
                                oForm.Freeze(false);
                                return;
                            }

                            SAPbobsCOM.CompanyService oCompanyService = null;
                            SAPbobsCOM.GeneralService oGeneralService = null;
                            SAPbobsCOM.GeneralData oGeneralData = null;
                            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                            oCompanyService = Program.oCompany.GetCompanyService();
                            oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXS_D");
                            oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                            oGeneralParams.SetProperty("DocEntry", docEntry);
                            oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                            string baseDocType;
                            switch (sCFL_ID)
                            {
                                case "Invoice_CFL": baseDocType = "ARInvoice"; break;
                                case "InvoiceCreditMemo_CFL": baseDocType = "ARCreditNote"; break;
                                case "InvoiceCorrectionInvoice_CFL": baseDocType = "ARCorrectionInvoice"; break;
                                case "DownPaymentInvoice_CFL": baseDocType = "ARDownPaymentRequest"; break;
                                case "ARDownPaymentVAT_CFL": baseDocType = "ARDownPaymentVAT"; break;
                                default: baseDocType = null; break;
                            }
                            List<string> exclList = new List<string>();

                            for (int i = 1; i <= oMatrix.RowCount; i++)
                            {
                                if (i != row && baseDocType == oMatrix.GetCellSpecific("U_baseDocT", i).Value.ToString() && oMatrix.GetCellSpecific("U_baseDTxt", i).Value.ToString() != "")
                                {
                                    exclList.Add(oMatrix.GetCellSpecific("U_baseDTxt", i).Value.ToString());
                                }
                            }
                            exclList.Add("0");

                            DataTable baseDocs = getListBaseDoc(oGeneralData, wbNumber, null, baseDocType, docEntry, exclList);

                            int docCount = baseDocs.Rows.Count;
                            SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
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
                            oCFL.SetConditions(oCons);
                        }
                        oForm.Freeze(false);
                    }
                }
                else if (beforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "BusinessPartner_CFL")
                        {
                            string businessPartnerCode = Convert.ToString(oDataTable.GetValue("CardCode", 0));
                            string businessPartnerName = Convert.ToString(oDataTable.GetValue("CardName", 0));
                            string businessPartnerTin = Convert.ToString(oDataTable.GetValue("LicTradNum", 0));

                            oForm.DataSources.DBDataSources.Item("@BDO_TAXS").SetValue("U_cardCode", 0, businessPartnerCode);
                            oForm.DataSources.DBDataSources.Item("@BDO_TAXS").SetValue("U_cardCodeN", 0, businessPartnerName);
                            oForm.DataSources.DBDataSources.Item("@BDO_TAXS").SetValue("U_cardCodeT", 0, businessPartnerTin);

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                        else if (sCFL_ID == "Invoice_CFL" || sCFL_ID == "InvoiceCreditMemo_CFL" || sCFL_ID == "InvoiceCorrectionInvoice_CFL" || sCFL_ID == "DownPaymentInvoice_CFL" || sCFL_ID == "ARDownPaymentVAT_CFL")
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wblMTR").Specific));
                            SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
                            if (cellPos == null)
                            {
                                return;
                            }

                            int docEntry = oDataTable.GetValue("DocEntry", 0);
                            double amount = 0; //თანხა დღგ-ის ჩათვლით
                            double amountTX = 0; //დღგ-ის თანხა
                            //double amtOutTX = 0; //თანხა დღგ-ის გარეშე
                            string wbNumber = null;
                            Dictionary<string, string> wblDocInfo = null;

                            if (sCFL_ID == "Invoice_CFL")
                            {
                                ARInvoice.getAmount(docEntry, out amount, out amountTX, out errorText);
                                wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(docEntry, "13", out errorText);
                                wbNumber = wblDocInfo["number"];
                            }
                            else if (sCFL_ID == "InvoiceCreditMemo_CFL")
                            {
                                ARCreditNote.getAmount(docEntry, out amount, out amountTX, out errorText);
                                wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(docEntry, "14", out errorText);
                                wbNumber = wblDocInfo["number"];
                            }
                            else if (sCFL_ID == "InvoiceCorrectionInvoice_CFL")
                            {
                                ArCorrectionInvoice.GetAmount(docEntry, out amount, out amountTX, out errorText);
                                wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(docEntry, "165", out errorText);
                                wbNumber = wblDocInfo["number"];
                            }
                            else if (sCFL_ID == "DownPaymentInvoice_CFL")
                            {
                                ARDownPaymentRequest.getAmount(docEntry, out amount, out amountTX);
                            }
                            else if (sCFL_ID == "ARDownPaymentVAT_CFL")
                            {
                                BDOSARDownPaymentVATAccrual.getAmount(docEntry, out amount, out amountTX);
                            }

                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("U_baseDTxt").Cells.Item(cellPos.rowIndex).Specific.Value = docEntry.ToString());
                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("U_baseDoc").Cells.Item(cellPos.rowIndex).Specific.Value = docEntry.ToString());
                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("U_amtBsDc").Cells.Item(cellPos.rowIndex).Specific.Value = amount.ToString(Nfi));
                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("U_tAmtBsDc").Cells.Item(cellPos.rowIndex).Specific.Value = amountTX.ToString(Nfi));
                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("U_wbNumber").Cells.Item(cellPos.rowIndex).Specific.Value = wbNumber);

                            CalculateAmount(oForm);

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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

        public static void CalculateAmount(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("wblMTR").Specific;
                oMatrix.FlushToDataSource();

                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDO_TXS1");

                int rowCount = oDBDataSourceMTR.Size - 1;

                decimal totalAmount = 0;
                decimal totalVAT = 0;

                for (int i = 0; i <= rowCount; i++)
                {
                    string baseDocT = oDBDataSourceMTR.GetValue("U_baseDocT", i);
                    if (!string.IsNullOrEmpty(baseDocT))
                    {
                        if (oDBDataSourceMTR.GetValue("U_baseDocT", i) == "ARCreditNote")
                        {
                            totalAmount = totalAmount - Convert.ToDecimal(oDBDataSourceMTR.GetValue("U_amtBsDc", i), CultureInfo.InvariantCulture);
                            totalVAT = totalVAT - Convert.ToDecimal(oDBDataSourceMTR.GetValue("U_tAmtBsDc", i), CultureInfo.InvariantCulture);
                        }

                        else if (oDBDataSourceMTR.GetValue("U_baseDocT", i) == "ARCorrectionInvoice")
                        {
                            totalAmount = totalAmount + Convert.ToDecimal(oDBDataSourceMTR.GetValue("U_amtBsDc", i), CultureInfo.InvariantCulture);
                            totalVAT = totalVAT + Convert.ToDecimal(oDBDataSourceMTR.GetValue("U_tAmtBsDc", i), CultureInfo.InvariantCulture);
                        }

                        else
                        {
                            totalAmount = totalAmount + Convert.ToDecimal(oDBDataSourceMTR.GetValue("U_amtBsDc", i), CultureInfo.InvariantCulture);
                            totalVAT = totalVAT + Convert.ToDecimal(oDBDataSourceMTR.GetValue("U_tAmtBsDc", i), CultureInfo.InvariantCulture);
                        }
                    }
                }
                //SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDO_TAXS");
                //oDBDataSource.SetValue("U_amount", 0, FormsB1.ConvertDecimalToString(totalAmount));
                //oDBDataSource.SetValue("U_amountTX", 0, FormsB1.ConvertDecimalToString(totalVAT));
                //oDBDataSource.SetValue("U_amtOutTX", 0, FormsB1.ConvertDecimalToString(totalAmount - totalVAT));

                LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("amountE").Specific.Value = FormsB1.ConvertDecimalToString(totalAmount));
                LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("amountTXE").Specific.Value = FormsB1.ConvertDecimalToString(totalVAT));
                LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("amtOutTXE").Specific.Value = FormsB1.ConvertDecimalToString(totalAmount - totalVAT));
            }
            catch (Exception ex)
            {
                Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static string getStatusValueByStatusNumber(string statusNumber, bool corrInv, bool refInv)
        {
            if (refInv)
                return "denied"; //უარყოფილი
            else if (statusNumber == "0")
                return "created"; //შექმნილი
            else if (statusNumber == "1")
                return "shipped"; //გადაგზავნილი
            else if (statusNumber == "2")
                return "confirmed"; //დადასტურებული
            else if (statusNumber == "-1")
                return "removed"; //წაშლილი
            else if (statusNumber == "9")
                return "incompleteShipped"; //არასრულად შექმნილი (9 შემოღებულია ჩვენს მიერ! (rs.ge-დან არ მოდის)) 
            //else if (statusNumber == "")
            //    return "paper"; //ქაღალდის
            //else if (statusNumber == "")
            //    return "disturbedSynchronization"; //დარღვეულია სინქრონიზაცია
            //else if (statusNumber == "")
            //    return "attachedToTheDeclaration"; //დეკლარაციაზე მიბმა
            else if (statusNumber == "6")
                return "cancellationProcess"; //გაუქმების პროცესში
            else if (statusNumber == "7")
                return "canceled"; //გაუქმებული           
            else if (statusNumber == "4")
                return "correctionCreated"; //შექმნილი კორექტირებული 
            else if (statusNumber == "5")
                return "correctionShipped"; //გადაგზავნილი კორექტირებული
            else if (statusNumber == "8")
                return "correctionConfirmed"; //დადასტურებული კორექტირებული
            else if (statusNumber == "3" & corrInv == false)
                return "primary"; //პირველადი
            else if (statusNumber == "3" & corrInv)
                return "corrected"; //კორექტირებული
            else
                return "empty";
        }

        public static Dictionary<string, object> getTaxInvoiceSentDocumentInfo(int docEntry, string baseDocType, string cardCode)
        {
            Dictionary<string, object> taxDocInfo = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT 
            ""BDO_TAXS"".""CreateDate"" AS ""createDate"",
            ""BDO_TAXS"".""U_opDate"" AS ""opDate"",
            ""BDO_TAXS"".""DocEntry"" AS ""docEntry"",
            ""BDO_TAXS"".""U_invID"" AS ""invID"",
            ""BDO_TAXS"".""U_number"" AS ""number"",
            ""BDO_TAXS"".""U_series"" AS ""series"",
            ""BDO_TAXS"".""U_status"" AS ""status"",
            ""BDO_TAXS"".""U_cardCodeT"" AS ""cardCodeT""
            FROM ""@BDO_TAXS"" AS ""BDO_TAXS"" 
            INNER JOIN ""@BDO_TXS1"" AS ""BDO_TXS1"" 
            ON ""BDO_TXS1"".""DocEntry"" = ""BDO_TAXS"".""DocEntry"" 
            WHERE ""BDO_TXS1"".""U_baseDoc"" = '" + docEntry + @"' AND ""BDO_TXS1"".""U_baseDocT"" = '" + baseDocType +
            @"' AND ""BDO_TAXS"".""U_cardCode"" = N'" + cardCode +
            @"' AND ""BDO_TAXS"".""Canceled"" = 'N' AND (""BDO_TAXS"".""U_status"" NOT IN ('removed', 'canceled') OR ""BDO_TAXS"".""U_status"" IS NULL)";

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
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
            return taxDocInfo;
        }

        public static void getPrimaryBaseDoc(int docEntry, string cardCode, out List<int> baseDocList)
        {
            int corrDoc = 0;
            string corrInv = null;
            baseDocList = new List<int>();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT 
            ""BDO_TAXS"".""CreateDate"" AS ""createDate"",
            ""BDO_TAXS"".""DocEntry"" AS ""docEntry"",
            ""BDO_TAXS"".""U_invID"" AS ""invID"",
            ""BDO_TAXS"".""U_number"" AS ""number"",
            ""BDO_TAXS"".""U_series"" AS ""series"",
            ""BDO_TAXS"".""U_status"" AS ""status"",
            ""BDO_TAXS"".""U_cardCodeT"" AS ""cardCodeT"",
            ""BDO_TAXS"".""U_corrInv"" AS ""corrInv"",            
            ""BDO_TAXS"".""U_corrDoc"" AS ""corrDoc"",             
            ""BDO_TXS1"".""U_baseDoc"" AS ""baseDoc"",
            ""BDO_TXS1"".""U_baseDocT"" AS ""baseDocT""
            FROM ""@BDO_TAXS"" AS ""BDO_TAXS"" 
            INNER JOIN ""@BDO_TXS1"" AS ""BDO_TXS1"" 
            ON ""BDO_TXS1"".""DocEntry"" = ""BDO_TAXS"".""DocEntry"" 
            WHERE           
            ""BDO_TAXS"".""U_cardCode"" = N'" + cardCode +
            @"' AND ""BDO_TAXS"".""DocEntry"" = '" + docEntry + "'";
            //@"' AND (""BDO_TAXS"".""Canceled"" = 'N' AND ""BDO_TAXS"".""U_status"" NOT IN ('removed', 'canceled'))";

            try
            {
                oRecordSet.DoQuery(query);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        corrInv = oRecordSet.Fields.Item("corrInv").Value.ToString();
                        //baseDocType = oRecordSet.Fields.Item("baseDocT").Value.ToString();

                        if (corrInv == "Y")
                            corrDoc = Convert.ToInt32(oRecordSet.Fields.Item("corrDoc").Value);
                        else
                            baseDocList.Add(Convert.ToInt32(oRecordSet.Fields.Item("baseDoc").Value));

                        oRecordSet.MoveNext();
                    }
                    if (corrInv == "Y" && corrDoc != 0)
                    {
                        getPrimaryBaseDoc(corrDoc, cardCode, out baseDocList);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
        }

        public static DataTable getListBaseDoc(SAPbobsCOM.GeneralData oGeneralData, string overhead_no, string overhead_dt_str, string baseDocType, int docEntryTaxInv, List<string> exclList)
        {
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
            string baseDocTable;
            string baseDocRowTable;
            string objectType;

            if (amount < 0)
            {
                amount = amount * (-1);
            }
            if (amountTX < 0)
            {
                amountTX = amountTX * (-1);
            }
            if (baseDocType == "ARInvoice") //A/R Invoice
            {
                baseDocTable = "OINV";
                baseDocRowTable = "INV1";
                objectType = "13";
            }
            else if (baseDocType == "ARCreditNote") //A/R CreditNote
            {
                baseDocTable = "ORIN";
                baseDocRowTable = "RIN1";
                objectType = "14";
            }
            else if (baseDocType == "ARCorrectionInvoice") //A/R Correction Invoice
            {
                baseDocTable = "OCSI";
                baseDocRowTable = "CSI1";
                objectType = "165";
            }
            else if (baseDocType == "ARDownPaymentRequest") //A/R DownPaymentInvoice
            {
                baseDocTable = "ODPI";
                baseDocRowTable = "DPI1";
                objectType = "203";
            }
            else if (baseDocType == "ARDownPaymentVAT") //A/R DownPaymentInvoice
            {
                baseDocTable = "@BDOSARDV";
                baseDocRowTable = "@BDOSRDV1";
                objectType = "UDO_F_BDO_ARDPV_D";
            }
            else
                return baseDocs;

            List<int> primaryBaseDocList = new List<int>();
            List<int> connectedDocList = new List<int>();
            if (corrInv)
            {
                getPrimaryBaseDoc(corrDoc, cardCode, out primaryBaseDocList);

                if (downPaymnt)
                {
                    connectedDocList = ARDownPaymentRequest.getAllConnectedDoc(primaryBaseDocList);
                }
                else
                {
                    connectedDocList = ARCreditNote.getAllConnectedDoc(primaryBaseDocList, "13");

                    if (baseDocTable == "OCSI")
                    {
                        connectedDocList = ArCorrectionInvoice.getAllConnectedDoc(primaryBaseDocList,"13");
                    }
                }
            }

            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string queryWbl = @"SELECT ""BDO_WBLD"".""U_baseDoc""
            FROM ""@BDO_WBLD"" AS ""BDO_WBLD"" 
            WHERE ""BDO_WBLD"".""Canceled"" = 'N'
            AND ""BDO_WBLD"".""U_baseDocT"" = '" + objectType + @"' 
            AND ""BDO_WBLD"".""U_baseDoc"" <> '0'
            AND ""BDO_WBLD"".""U_cardCode"" = N'" + cardCode + @"'
            AND ""BDO_WBLD"".""U_number"" = '" + overhead_no + @"'";

            string query = @"SELECT
	             ""TABL"".""DocEntry"",
	             ""TABL"".""DocDate"" as ""DocDate"",
	             '" + baseDocType + @"' AS ""BaseDocType"",
	             SUM(""TBL1"".""GTotal"") AS ""GTotal"",
            	 SUM(""TBL1"".""LineVat"") AS ""LineVat"" 
            FROM """ + baseDocTable + @""" AS ""TABL"" 
            LEFT JOIN """ + baseDocRowTable + @""" AS ""TBL1"" ON ""TBL1"".""DocEntry"" = ""TABL"".""DocEntry"" 
            WHERE ""TABL"".""CANCELED"" = 'N' 
            AND ""TABL"".""CardCode"" = N'" + cardCode + @"'
            AND ""TABL"".""DocDate"" >= '" + firstDayMonth.ToString("yyyyMMdd") + @"' AND ""TABL"".""DocDate"" <= '" + lastDayMonth.ToString("yyyyMMdd") + "'";

            if (baseDocType != "ARDownPaymentRequest" && string.IsNullOrEmpty(overhead_no) == false) //ვეძებთ ზედნადების ნომრით
            {
                query = query + @" AND ""TABL"".""DocEntry"" IN (" + queryWbl + ")";
            }
            if (baseDocType == "ARDownPaymentRequest")
            {
                query = query.Replace(@"""TBL1"".""GTotal""", @"""TBL1"".""U_BDOSDPMAmt""");
                query = query.Replace(@"""TBL1"".""LineVat""", @"""TBL1"".""U_BDOSDPMVat""");
                //query = query + @" AND ""TABL"".""Posted"" = 'Y'";
            }
            if (corrInv)
            {
                query = query + @" AND ""TABL"".""DocEntry"" IN (" + string.Join(",", connectedDocList) + ")";
            }
            query = query + @" AND ""TABL"".""DocEntry"" NOT IN ( SELECT
            	 ""BDO_TXS1"".""U_baseDoc"" 
            	FROM ""@BDO_TAXS"" AS ""BDO_TAXS"" 
            	INNER JOIN ""@BDO_TXS1"" AS ""BDO_TXS1"" ON ""BDO_TXS1"".""DocEntry"" = ""BDO_TAXS"".""DocEntry"" 
            	WHERE ""BDO_TAXS"".""Canceled"" = 'N' 
            		AND (""BDO_TAXS"".""U_status"" NOT IN ('removed',
            	 'canceled') OR ""BDO_TAXS"".""U_status"" IS NULL) 
            	AND ""BDO_TXS1"".""U_baseDocT"" = '" + baseDocType + @"' 
            	AND ""BDO_TAXS"".""U_cardCode"" = N'" + cardCode + @"'  
            	AND ""BDO_TXS1"".""DocEntry"" <> '" + docEntryTaxInv + @"' ) 
                AND ""TABL"".""DocEntry"" NOT IN (" + string.Join(",", exclList) + @")
            GROUP BY ""TABL"".""DocEntry"",
            	 ""TABL"".""DocDate""";

            if (baseDocType == "ARDownPaymentVAT")
            {
                query = query.Replace(@"""TABL"".""CANCELED""", @"""TABL"".""Canceled""");
                query = query.Replace(@"""TABL"".""DocDate""", @"""TABL"".""U_DocDate""");
                query = query.Replace(@"""TABL"".""CardCode""", @"""TABL"".""U_cardCode""");
                ;

                query = query.Replace(@"""TBL1"".""GTotal""", @"""TBL1"".""U_GrsAmnt""");
                query = query.Replace(@"""TBL1"".""LineVat""", @"""TBL1"".""U_VatAmount""");
                //query = query + @" AND ""TABL"".""Posted"" = 'Y'";
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
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
            return baseDocs;
        }

        public static string getAllConnectedDoc(int docEntry, ref List<int> docEntryARInvoiceList,
            ref List<int> docEntryARCreditNoteList, ref List<int> docEntryARCorrectionInvoiceList)
        {
            int corrDoc = 0;
            string corrInv = null;
            string corrDocStr = null;
            string baseDocStr = null;

            SAPbobsCOM.Recordset oRecordSet =
                (SAPbobsCOM.Recordset) Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT 
            ""BDO_TAXS"".""CreateDate"" AS ""createDate"",
            ""BDO_TAXS"".""DocEntry"" AS ""docEntry"",
            ""BDO_TAXS"".""U_invID"" AS ""invID"",
            ""BDO_TAXS"".""U_number"" AS ""number"",
            ""BDO_TAXS"".""U_series"" AS ""series"",
            ""BDO_TAXS"".""U_status"" AS ""status"",
            ""BDO_TAXS"".""U_cardCodeT"" AS ""cardCodeT"",
            ""BDO_TAXS"".""U_corrInv"" AS ""corrInv"",            
            ""BDO_TAXS"".""U_corrDoc"" AS ""corrDoc"",             
            ""BDO_TXS1"".""U_baseDoc"" AS ""baseDoc"",
            ""BDO_TXS1"".""U_baseDocT"" AS ""baseDocT""
            FROM ""@BDO_TAXS"" AS ""BDO_TAXS"" 
            INNER JOIN ""@BDO_TXS1"" AS ""BDO_TXS1"" 
            ON ""BDO_TXS1"".""DocEntry"" = ""BDO_TAXS"".""DocEntry"" 
            WHERE ""BDO_TAXS"".""DocEntry"" = '" + docEntry + @"'
            AND (""BDO_TAXS"".""Canceled"" = 'N' AND ""BDO_TAXS"".""U_status"" NOT IN ('removed', 'canceled'))";

            string baseDocT;
            try
            {
                oRecordSet.DoQuery(query);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        corrInv = oRecordSet.Fields.Item("corrInv").Value.ToString();
                        corrDocStr = oRecordSet.Fields.Item("corrDoc").Value.ToString();
                        if (corrInv == "Y" && string.IsNullOrEmpty(corrDocStr) == false)
                            corrDoc = Convert.ToInt32(corrDocStr);

                        baseDocT = oRecordSet.Fields.Item("baseDocT").Value.ToString();
                        baseDocStr = oRecordSet.Fields.Item("baseDoc").Value.ToString();
                        if (string.IsNullOrEmpty(baseDocStr) == false)
                        {
                            if (baseDocT == "ARInvoice") //A/R Invoice
                            {
                                docEntryARInvoiceList.Add(Convert.ToInt32(baseDocStr));
                            }
                            else if (baseDocT == "ARCreditNote") //A/R Credit Note
                            {
                                docEntryARCreditNoteList.Add(Convert.ToInt32(baseDocStr));
                            }
                            else if (baseDocT == "ARCorrectionInvoice") //A/R Correction Invoice
                            {
                                docEntryARCorrectionInvoiceList.Add(Convert.ToInt32(baseDocStr));
                            }
                        }

                        oRecordSet.MoveNext();
                    }

                    if (corrInv == "Y" && corrDoc != 0)
                    {
                        getAllConnectedDoc(corrDoc, ref docEntryARInvoiceList, ref docEntryARCreditNoteList,
                            ref docEntryARCorrectionInvoiceList);
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        public static void addMatrixRow(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("wblMTR").Specific;
                oMatrix.FlushToDataSource();

                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDO_TXS1");
                if (!string.IsNullOrEmpty(oDBDataSourceMTR.GetValue("U_baseDocT", oDBDataSourceMTR.Size - 1)))
                {
                    oDBDataSourceMTR.InsertRecord(oDBDataSourceMTR.Size);
                }
                oDBDataSourceMTR.SetValue("LineId", oDBDataSourceMTR.Size - 1, oDBDataSourceMTR.Size.ToString());

                string downPaymnt = oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("U_downPaymnt", 0).Trim();
                if (downPaymnt == "Y")
                {
                    oDBDataSourceMTR.SetValue("U_baseDocT", oDBDataSourceMTR.Size - 1, "ARDownPaymentRequest");
                }
                else
                {
                    oDBDataSourceMTR.SetValue("U_baseDocT", oDBDataSourceMTR.Size - 1, "ARInvoice");
                }

                oMatrix.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void deleteMatrixRow(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
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
                        oForm.DataSources.DBDataSources.Item("@BDO_TXS1").RemoveRecord(row - deletedRowCount);
                        firstRow = row;
                    }
                }

                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDO_TXS1");
                int rowCount = oDBDataSourceMTR.Size;

                for (int i = 1; i <= rowCount; i++)
                {
                    string itemCode = oDBDataSourceMTR.GetValue("U_baseDocT", i - 1);
                    if (!string.IsNullOrEmpty(itemCode))
                    {
                        oDBDataSourceMTR.SetValue("LineId", i - 1, i.ToString());
                    }
                }
                oMatrix.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && deletedRowCount > 0)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                GC.Collect();
                oForm.Freeze(false);
            }
        }

        public static void formDataAddUpdate(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                string statusInv = oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("U_status", 0).Trim();
                if (statusInv == "confirmed" || statusInv == "correctionConfirmed")
                {
                    errorText = BDOSResources.getTranslate("UpdateConfirmedTaxInvoiceNotAllowed") + "!";
                    return;
                }

                //
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDO_TXS1");

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

        public static void createDocument(string objectType, int baseDocEntry, string corrType, bool fromDoc, int answer, SAPbobsCOM.GeneralData oGeneralData, bool union, List<int> docEntryARCreditNoteList, List<int> docEntryARCorrectionInvoiceList, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;

            if (objectType == "13")  //A/R Invoice
            {
                createDocumentARInvoiceType(baseDocEntry, fromDoc, answer, oGeneralData, union, docEntryARCreditNoteList, docEntryARCorrectionInvoiceList, out newDocEntry, out errorText);
            }
            else if (objectType == "14") //A/R Credit Memo 
            {
                createDocumentInvoiceCreditMemoType(baseDocEntry, corrType, fromDoc, answer, oGeneralData, union, out newDocEntry, out errorText);
            }
            else if (objectType == "165") //A/R Correction Invoice
            {
                createDocumentInvoiceCorrectionInvoiceType(baseDocEntry, corrType, fromDoc, answer, oGeneralData, union, out newDocEntry, out errorText);
            }
            else if (objectType == "203") //A/R Down Payment Invoice 
            {
                createDocumentInvoiceARDownPaymentRequestType(baseDocEntry, corrType, fromDoc, out newDocEntry, out errorText);
            }
            else if (objectType == "UDO_F_BDO_ARDPV_D") //A/R Down Payment VAT 
            {
                createDocumentInvoiceARDownPaymentVAT(baseDocEntry, corrType, fromDoc, out newDocEntry, out errorText);
            }
        }

        public static void createDocumentForUnion(DataTable UnionTable, string PrevCardCode, DateTime PrevOpDate, ref string docEntry, out string errorText)
        {
            errorText = null;

            createDocumentARInvoiceTypeForUnion(PrevCardCode, PrevOpDate, UnionTable, ref docEntry, out errorText);
        }

        private static void createDocumentARInvoiceType(int baseDocEntry, bool fromDoc, int answer, SAPbobsCOM.GeneralData oGeneralData, bool union, List<int> docEntryARCreditNoteList, List<int> docEntryARCorrectionInvoiceList, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "SELECT " +
            "\"OINV\".\"DocEntry\", " +
            "\"OINV\".\"DocDate\", " +
            "\"OINV\".\"CardCode\", " +
            "\"OCRD\".\"CardName\", " +
            "\"OCRD\".\"LicTradNum\", " +
            "SUM(\"INV1\".\"GTotal\") as \"DocTotal\", " +
            "SUM(\"INV1\".\"LineVat\") as \"VatSum\"  " +
            "FROM \"INV1\"  " +
            "left join \"OINV\" on \"INV1\".\"DocEntry\" = \"OINV\".\"DocEntry\"  " +
            "left join \"OCRD\" on \"OINV\".\"CardCode\" = \"OCRD\".\"CardCode\"  " +
            "WHERE \"INV1\".\"DocEntry\" = '" + baseDocEntry + "'  " +
            "GROUP BY \"OINV\".\"DocEntry\", " +
            "\"OINV\".\"DocDate\", " +
            "\"OINV\".\"CardCode\", " +
            "\"OCRD\".\"CardName\", " +
            "\"OCRD\".\"LicTradNum\"  ";

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    decimal amount = Convert.ToDecimal(oRecordSet.Fields.Item("DocTotal").Value);
                    decimal amountTX = Convert.ToDecimal(oRecordSet.Fields.Item("VatSum").Value);
                    decimal amtOutTX = amount - amountTX;
                    DateTime docDate = oRecordSet.Fields.Item("DocDate").Value;
                    string cardCode = oRecordSet.Fields.Item("CardCode").Value.ToString();

                    SAPbobsCOM.CompanyService oCompanyService = null;
                    SAPbobsCOM.GeneralService oGeneralService = null;

                    if (union == false)
                    {
                        oCompanyService = Program.oCompany.GetCompanyService();
                        oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXS_D");
                        oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                        oGeneralData.SetProperty("U_opDate", new DateTime(docDate.Year, docDate.Month, 1));
                        oGeneralData.SetProperty("U_cardCode", cardCode);
                        oGeneralData.SetProperty("U_cardCodeN", oRecordSet.Fields.Item("CardName").Value.ToString());
                        oGeneralData.SetProperty("U_cardCodeT", oRecordSet.Fields.Item("LicTradNum").Value.ToString());
                        oGeneralData.SetProperty("U_status", "empty");
                    }

                    Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(baseDocEntry, "13", out errorText);

                    SAPbobsCOM.GeneralDataCollection oChildren = null;

                    oChildren = oGeneralData.Child("BDO_TXS1");

                    SAPbobsCOM.GeneralData oChild = oChildren.Add();

                    oChild.SetProperty("U_baseDocT", "ARInvoice"); //რეალიზაცია
                    oChild.SetProperty("U_baseDoc", oRecordSet.Fields.Item("DocEntry").Value);
                    oChild.SetProperty("U_baseDTxt", oRecordSet.Fields.Item("DocEntry").Value.ToString());
                    oChild.SetProperty("U_amtBsDc", oRecordSet.Fields.Item("DocTotal").Value); //თანხა დღგ-ის ჩათვლით
                    oChild.SetProperty("U_tAmtBsDc", oRecordSet.Fields.Item("VatSum").Value); //დღგ-ის თანხა
                    oChild.SetProperty("U_wbNumber", wblDocInfo["number"]);

                    
                    List<int> connectedDocList = ARInvoice.getAllConnectedDoc(new List<int>() { baseDocEntry }, "13", docDate, new DateTime(), 0, out errorText);
                    int rowCountCN = connectedDocList.Count();

                    List<int> connectedDocListCI = ARInvoice.getAllConnectedARCorrectionDoc(new List<int>() { baseDocEntry }, "13", docDate, new DateTime(), 0, out errorText);
                    int rowCount = connectedDocListCI.Count();

                    if (rowCount != 0 || rowCountCN != 0)
                    {
                        if (fromDoc)
                        {
                            answer = Program.uiApp.MessageBox(
                                BDOSResources.getTranslate(
                                    "ARCDocumentsOnARIDoYouWantToCreateUnitedTaxInvoiceIncludingTheseDocuments") + "?",
                                1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"),
                                ""); //არსებობს რეალიზაციის დოკუმენტზე რეალიზაციის კორექტირების დოკუმენტები, გსურთ შეიქმნას ერთიანი ფაქტურა ამ დოკუმენტების გათვალისწინებით                   
                        }

                        if (answer == 1 || union)
                        {
                            //კორექტირების დოკუმენტები (AR Credit Note)
                            if (rowCountCN != 0)
                            {
                                wblDocInfo = null;
                                int corrDocEntry = 0;
                                string corrInvID = null;
                                string BDO_CNTp = null;
                                int creditNoteDocEntry = 0;

                                for (int i = 0; i < rowCountCN; i++)
                                {
                                    creditNoteDocEntry = connectedDocList[i];
                                    if (union && docEntryARCreditNoteList.Contains(creditNoteDocEntry) == false)
                                    {
                                        continue;
                                    }

                                    BDO_CNTp = CommonFunctions.getValue("ORIN", "U_BDO_CNTp", "DocEntry",
                                        creditNoteDocEntry.ToString()).ToString(); //0=კორექტირება, 1=დაბრუნება

                                    Dictionary<string, object> taxDocInfo =
                                        getTaxInvoiceSentDocumentInfo(creditNoteDocEntry, "ARCreditNote", cardCode);
                                    if (taxDocInfo != null)
                                    {
                                        corrDocEntry =
                                            Convert.ToInt32(
                                                taxDocInfo[
                                                    "docEntry"]); //კორექტირების ა/ფ-ის Entry //წესით არ უნდა იყოს შევსებული
                                        corrInvID = taxDocInfo["invID"]
                                            .ToString(); //კორექტირების ა/ფ-ის ID //წესით არ უნდა იყოს შევსებული
                                    }

                                    if (corrDocEntry == 0 && string.IsNullOrEmpty(corrInvID))
                                    {
                                        double gTotal;
                                        double lineVat;
                                        ARCreditNote.getAmount(creditNoteDocEntry, out gTotal, out lineVat,
                                            out errorText);

                                        if (BDO_CNTp == "0") //კორექტირება
                                        {
                                            wblDocInfo =
                                                BDO_Waybills.getWaybillDocumentInfo(baseDocEntry, "13", out errorText);
                                        }
                                        else //დაბრუნება
                                        {
                                            wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(creditNoteDocEntry, "14",
                                                out errorText);
                                        }

                                        oChild = oChildren.Add();
                                        oChild.SetProperty("U_baseDocT", "ARCreditNote");
                                        oChild.SetProperty("U_baseDoc", creditNoteDocEntry);
                                        oChild.SetProperty("U_baseDTxt", creditNoteDocEntry.ToString());
                                        oChild.SetProperty("U_amtBsDc", gTotal); //თანხა დღგ-ის ჩათვლით
                                        oChild.SetProperty("U_tAmtBsDc", lineVat); //დღგ-ის თანხა
                                        oChild.SetProperty("U_wbNumber", wblDocInfo["number"]);

                                        amount = amount - Convert.ToDecimal(gTotal);
                                        amountTX = amountTX - Convert.ToDecimal(lineVat);
                                    }
                                    else
                                    {
                                        errorText = BDOSResources.getTranslate(
                                                        "TaxInvoiceIsAlreadyCreatedBasedOnARCorrection") + "! ID : " +
                                                    corrInvID;
                                        return;
                                    }
                                }
                            }

                            //კორექტირების დოკუმენტები (AR Correction Invoice)
                            if (rowCount != 0)
                            {
                                wblDocInfo = null;
                                int corrDocEntry = 0;
                                string corrInvID = null;
                                string BDOSCITp = null;
                                int correctionInvoiceDocEntry = 0;

                                for (int i = 0; i < rowCount; i++)
                                {
                                    correctionInvoiceDocEntry = connectedDocListCI[i];
                                    if (union && docEntryARCorrectionInvoiceList.Contains(correctionInvoiceDocEntry) ==
                                        false)
                                    {
                                        continue;
                                    }

                                    BDOSCITp = CommonFunctions
                                        .getValue("OCSI", "U_BDOSCITp", "DocEntry",
                                            correctionInvoiceDocEntry.ToString())
                                        .ToString(); //0=კორექტირება, 1=დაბრუნება

                                    Dictionary<string, object> taxDocInfo =
                                        getTaxInvoiceSentDocumentInfo(correctionInvoiceDocEntry, "ARCorrectionInvoice",
                                            cardCode);
                                    if (taxDocInfo != null)
                                    {
                                        corrDocEntry =
                                            Convert.ToInt32(
                                                taxDocInfo[
                                                    "docEntry"]); //კორექტირების ა/ფ-ის Entry //წესით არ უნდა იყოს შევსებული
                                        corrInvID = taxDocInfo["invID"]
                                            .ToString(); //კორექტირების ა/ფ-ის ID //წესით არ უნდა იყოს შევსებული
                                    }

                                    if (corrDocEntry == 0 && string.IsNullOrEmpty(corrInvID))
                                    {
                                        double gTotal;
                                        double lineVat;
                                        ArCorrectionInvoice.GetAmount(correctionInvoiceDocEntry, out gTotal,
                                            out lineVat,
                                            out errorText);

                                        if (BDOSCITp == "0") //კორექტირება
                                        {
                                            wblDocInfo =
                                                BDO_Waybills.getWaybillDocumentInfo(baseDocEntry, "13", out errorText);
                                        }
                                        else //დაბრუნება
                                        {
                                            wblDocInfo =
                                                BDO_Waybills.getWaybillDocumentInfo(correctionInvoiceDocEntry, "165",
                                                    out errorText);
                                        }

                                        oChild = oChildren.Add();
                                        oChild.SetProperty("U_baseDocT", "ARCorrectionInvoice");
                                        oChild.SetProperty("U_baseDoc", correctionInvoiceDocEntry);
                                        oChild.SetProperty("U_baseDTxt", correctionInvoiceDocEntry.ToString());
                                        oChild.SetProperty("U_amtBsDc", gTotal); //თანხა დღგ-ის ჩათვლით
                                        oChild.SetProperty("U_tAmtBsDc", lineVat); //დღგ-ის თანხა
                                        oChild.SetProperty("U_wbNumber", wblDocInfo["number"]);

                                        amount = amount + Convert.ToDecimal(gTotal);
                                        amountTX = amountTX + Convert.ToDecimal(lineVat);
                                    }
                                    else
                                    {
                                        errorText =
                                            BDOSResources.getTranslate(
                                                "TaxInvoiceIsAlreadyCreatedBasedOnARCorrection") +
                                            "! ID : " + corrInvID;
                                        return;
                                    }
                                }
                            }
                        }
                    }



                    if (union == false)
                    {
                        oGeneralData.SetProperty("U_amount", Convert.ToDouble(amount)); //თანხა დღგ-ის ჩათვლით
                        oGeneralData.SetProperty("U_amountTX", Convert.ToDouble(amountTX)); //დღგ-ის თანხა
                        oGeneralData.SetProperty("U_amtOutTX", Convert.ToDouble(amount - amountTX)); //თანხა დღგ-ის გარეშე

                        try
                        {
                            var response = oGeneralService.Add(oGeneralData);
                            var docEntry = response.GetProperty("DocEntry");
                            newDocEntry = Convert.ToInt32(docEntry);
                        }
                        catch (Exception ex)
                        {
                            errorText = BDOSResources.getTranslate("ErrorDocumentAddEdit") + " : " + ex.Message;
                        }
                    }
                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        public static SAPbobsCOM.GeneralData createDocumentForUnion(string cardCode, DateTime docDate, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;

            oCompanyService = Program.oCompany.GetCompanyService();
            oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXS_D");
            oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

            try
            {
                SAPbobsCOM.BusinessPartners oBP;
                oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                if (oBP.GetByKey(cardCode))
                {
                    oGeneralData.SetProperty("U_opDate", new DateTime(docDate.Year, docDate.Month, 1));
                    oGeneralData.SetProperty("U_cardCode", cardCode);
                    oGeneralData.SetProperty("U_cardCodeN", oBP.UserFields.Fields.Item("CardName").Value.ToString());
                    oGeneralData.SetProperty("U_cardCodeT", oBP.UserFields.Fields.Item("LicTradNum").Value.ToString());
                    oGeneralData.SetProperty("U_status", "empty");
                }
                else
                {
                    errorText = BDOSResources.getTranslate("CouldNotFindBusinessPartner") + " : " + cardCode;
                    Marshal.FinalReleaseComObject(oGeneralService);
                    Marshal.FinalReleaseComObject(oBP);
                    return null;
                }
                return oGeneralData;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return null;
            }
            finally
            {
                Marshal.ReleaseComObject(oGeneralService);
            }
        }

        private static void createDocumentARInvoiceTypeForUnion(string cardCode, DateTime docDate, DataTable baseDocs, ref string docEntry, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;

            oCompanyService = Program.oCompany.GetCompanyService();
            oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXS_D");
            oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

            SAPbobsCOM.BusinessPartners oBP;
            oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
            if (oBP.GetByKey(cardCode))
            {
                oGeneralData.SetProperty("U_opDate", new DateTime(docDate.Year, docDate.Month, 1));
                oGeneralData.SetProperty("U_cardCode", cardCode);
                oGeneralData.SetProperty("U_cardCodeN", oBP.UserFields.Fields.Item("CardName").Value.ToString());
                oGeneralData.SetProperty("U_cardCodeT", oBP.UserFields.Fields.Item("LicTradNum").Value.ToString());
                oGeneralData.SetProperty("U_status", "empty");
            }
            else
            {
                errorText = BDOSResources.getTranslate("CouldNotFindBusinessPartner") + " : " + cardCode;
                Marshal.FinalReleaseComObject(oGeneralService);
                Marshal.FinalReleaseComObject(oBP);
                return;
            }

            SAPbobsCOM.GeneralDataCollection oChildren = oGeneralData.Child("BDO_TXS1");
            SAPbobsCOM.GeneralData oChild = null;

            Dictionary<string, string> wblDocInfo = null;

            decimal amount = 0;
            decimal amountTX = 0;
            string baseDocType = null;
            string status = null;
            int InvoiceEntry = 0;
            int CreditMemoEntry = 0;
            int ARInvoiceDocEntry = 0;
            bool UnionTableWithInvoice = false;
            string baseDocTWb = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            int rowCount = baseDocs.Rows.Count;

            for (int i = 0; i < rowCount; i++)
            {
                baseDocType = "ARInvoice";
                DataRow taxDataRow = baseDocs.Rows[i];

                InvoiceEntry = 0;
                CreditMemoEntry = 0;

                try
                {
                    InvoiceEntry = Convert.ToInt32(taxDataRow["InvEntry"]);
                    UnionTableWithInvoice = true;
                }
                catch
                {
                    InvoiceEntry = 0;
                }

                int baseDoc = InvoiceEntry;

                try
                {
                    CreditMemoEntry = Convert.ToInt32(taxDataRow["CrMEntry"]);
                }
                catch
                {
                    CreditMemoEntry = 0;
                }

                if (baseDoc == 0)
                {
                    baseDoc = CreditMemoEntry;
                }

                if (InvoiceEntry > 0)
                {
                    baseDocType = "ARInvoice";
                    baseDocTWb = "13";
                    ARInvoiceDocEntry = InvoiceEntry;
                }

                else if (CreditMemoEntry > 0)
                {
                    baseDocType = "ARCreditNote";
                    string query = "select  \"U_BDO_CNTp\" from \"ORIN\" Where \"DocEntry\" = " + CreditMemoEntry;

                    oRecordSet.DoQuery(query);

                    if (!oRecordSet.EoF)
                    {
                        string U_BDO_CNTp = oRecordSet.Fields.Item("U_BDO_CNTp").Value;

                        if (U_BDO_CNTp == "0") //კორექტირება
                            baseDocTWb = "13";
                        else
                            baseDocTWb = "14";
                    }
                    ARCreditNote.getBaseDoc(baseDoc, "13", out ARInvoiceDocEntry);
                }

                if (CreditMemoEntry > 0)
                {
                    DataRow[] foundrows = baseDocs.Select("InvEntry = '" + ARInvoiceDocEntry + "'");
                    if (foundrows.Length == 0)
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("NoBaseSaleFoundTaxInvoiceUnion") + taxDataRow["docNum"]);
                        continue;
                    }
                }

                if (baseDocTWb == "13")
                {
                    wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(ARInvoiceDocEntry, "13", out errorText);
                }
                else
                {
                    wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(baseDoc, baseDocTWb, out errorText);
                }

                string wblNum = wblDocInfo["number"];
                status = wblDocInfo["status"];

                oChild = oChildren.Add();
                oChild.SetProperty("U_baseDocT", baseDocType);
                oChild.SetProperty("U_baseDoc", Convert.ToInt32(baseDoc));
                oChild.SetProperty("U_baseDTxt", baseDoc.ToString());
                oChild.SetProperty("U_amtBsDc", Convert.ToDouble(taxDataRow["Sum"], CultureInfo.InvariantCulture)); //თანხა დღგ-ის ჩათვლით
                oChild.SetProperty("U_tAmtBsDc", Convert.ToDouble(taxDataRow["VatSum"], CultureInfo.InvariantCulture)); //დღგ-ის თანხა
                oChild.SetProperty("U_wbNumber", wblNum);

                if (baseDocType == "ARInvoice")
                {
                    amount = amount + Convert.ToDecimal(Convert.ToDouble(taxDataRow["Sum"], CultureInfo.InvariantCulture));
                    amountTX = amountTX + Convert.ToDecimal(Convert.ToDouble(taxDataRow["VatSum"], CultureInfo.InvariantCulture));
                }
                else
                {
                    amount = amount - Convert.ToDecimal(Convert.ToDouble(taxDataRow["Sum"], CultureInfo.InvariantCulture));
                    amountTX = amountTX - Convert.ToDecimal(Convert.ToDouble(taxDataRow["VatSum"], CultureInfo.InvariantCulture));
                }
            }

            if (UnionTableWithInvoice == false)
            {
                errorText = BDOSResources.getTranslate("UnionTableMustContainAtLeastOneInvoice");
                return;
            }

            oGeneralData.SetProperty("U_amount", Convert.ToDouble(amount)); //თანხა დღგ-ის ჩათვლით
            oGeneralData.SetProperty("U_amountTX", Convert.ToDouble(amountTX)); //დღგ-ის თანხა
            oGeneralData.SetProperty("U_amtOutTX", Convert.ToDouble(amount - amountTX)); //თანხა დღგ-ის გარეშე

            try
            {
                var response = oGeneralService.Add(oGeneralData);
                docEntry = Convert.ToString(response.GetProperty("DocEntry"));
            }
            catch (Exception ex)
            {
                errorText = BDOSResources.getTranslate("ErrorDocumentAddEdit") + " : " + ex.Message;
            }
        }

        private static void createDocumentInvoiceCreditMemoType(int baseDocEntry, string corrType, bool fromDoc, int answer, SAPbobsCOM.GeneralData oGeneralData, bool union, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "SELECT " +
            "\"ORIN\".\"DocEntry\", " +
            "\"ORIN\".\"DocDate\", " +
            "\"ORIN\".\"DocTime\", " +
            "\"ORIN\".\"CardCode\", " +
            "\"ORIN\".\"U_BDO_CNTp\", " +
            "\"OCRD\".\"CardName\", " +
            "\"OCRD\".\"LicTradNum\", " +
            "SUM(\"RIN1\".\"GTotal\") as \"DocTotal\", " +
            "SUM(\"RIN1\".\"LineVat\") as \"VatSum\"  " +
            "FROM \"RIN1\"  " +
            "left join \"ORIN\" on \"RIN1\".\"DocEntry\" = \"ORIN\".\"DocEntry\"  " +
            "left join \"OCRD\" on \"ORIN\".\"CardCode\" = \"OCRD\".\"CardCode\"  " +
            "WHERE \"RIN1\".\"DocEntry\" = '" + baseDocEntry + "'  " +
            "GROUP BY \"ORIN\".\"DocEntry\", " +
            "\"ORIN\".\"DocDate\", " +
            "\"ORIN\".\"DocTime\", " +
            "\"ORIN\".\"CardCode\", " +
            "\"ORIN\".\"U_BDO_CNTp\", " +
            "\"OCRD\".\"CardName\", " +
            "\"OCRD\".\"LicTradNum\"  ";

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    decimal amount = Convert.ToDecimal(oRecordSet.Fields.Item("DocTotal").Value);
                    decimal amountTX = Convert.ToDecimal(oRecordSet.Fields.Item("VatSum").Value);
                    decimal amtOutTX = amount - amountTX;

                    DateTime docDate = oRecordSet.Fields.Item("DocDate").Value;
                    int docTime = oRecordSet.Fields.Item("DocTime").Value;
                    string cardCode = oRecordSet.Fields.Item("CardCode").Value.ToString();
                    string BDO_CNTp = oRecordSet.Fields.Item("U_BDO_CNTp").Value.ToString();

                    SAPbobsCOM.CompanyService oCompanyService = null;
                    SAPbobsCOM.GeneralService oGeneralService = null;

                    if (union == false)
                    {
                        oCompanyService = Program.oCompany.GetCompanyService();
                        oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXS_D");
                        oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                        oGeneralData.SetProperty("U_opDate", new DateTime(docDate.Year, docDate.Month, 1));
                        oGeneralData.SetProperty("U_cardCode", cardCode);
                        oGeneralData.SetProperty("U_cardCodeN", oRecordSet.Fields.Item("CardName").Value.ToString());
                        oGeneralData.SetProperty("U_cardCodeT", oRecordSet.Fields.Item("LicTradNum").Value.ToString());
                        oGeneralData.SetProperty("U_status", "empty");
                    }

                    int ARInvoiceDocEntry = 0;
                    ARCreditNote.getBaseDoc(baseDocEntry, "13", out ARInvoiceDocEntry);

                    Dictionary<string, string> wblDocInfo = null;
                    if (BDO_CNTp == "0" && ARInvoiceDocEntry != 0) //კორექტირება                   
                        wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(ARInvoiceDocEntry, "13", out errorText);
                    else
                        wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(baseDocEntry, "14", out errorText);

                    if (ARInvoiceDocEntry == 0)
                    {
                        if (union == false)
                        {
                            errorText = BDOSResources.getTranslate("BaseARInvoiceCouldNotFound") + "! ";
                            return;
                        }

                        if (fromDoc)
                        {
                            answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("BaseARInvoiceCouldNotFound") + "! " + BDOSResources.getTranslate("ContinueCreatingTaxInvoice") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                        }

                        if (answer == 1)
                        {
                            try
                            {
                                //საფუძველის ცხრილის შევსება
                                SAPbobsCOM.GeneralDataCollection oChildren = null;
                                oChildren = oGeneralData.Child("BDO_TXS1");

                                SAPbobsCOM.GeneralData oChild = oChildren.Add();

                                oChild.SetProperty("U_baseDocT", "ARCreditNote"); //A/R CreditNote
                                oChild.SetProperty("U_baseDoc", baseDocEntry);
                                oChild.SetProperty("U_baseDTxt", baseDocEntry.ToString());
                                oChild.SetProperty("U_amtBsDc", oRecordSet.Fields.Item("DocTotal").Value); //თანხა დღგ-ის ჩათვლით
                                oChild.SetProperty("U_tAmtBsDc", oRecordSet.Fields.Item("VatSum").Value); //დღგ-ის თანხა
                                oChild.SetProperty("U_wbNumber", wblDocInfo["number"]);

                                if (union == false)
                                {
                                    oGeneralData.SetProperty("U_amount", Convert.ToDouble(amount)); //თანხა დღგ-ის ჩათვლით
                                    oGeneralData.SetProperty("U_amountTX", Convert.ToDouble(amountTX)); //დღგ-ის თანხა
                                    oGeneralData.SetProperty("U_amtOutTX", Convert.ToDouble(amount - amountTX)); //თანხა დღგ-ის გარეშე

                                    var response = oGeneralService.Add(oGeneralData);
                                    var docEntry = response.GetProperty("DocEntry");
                                    newDocEntry = Convert.ToInt32(docEntry);
                                }
                            }
                            catch (Exception ex)
                            {
                                errorText = BDOSResources.getTranslate("ErrorDocumentAddEdit") + " : " + ex.Message;
                                return;
                            }
                        }
                        return;
                    }
                    else if (union)
                    {
                        SAPbobsCOM.GeneralDataCollection oChildren = null;
                        oChildren = oGeneralData.Child("BDO_TXS1");
                        int count = oChildren.Count;
                        if (count > 0)
                        {
                            for (int i = 0; i < count; i++)
                            {
                                errorText = BDOSResources.getTranslate("BaseARInvoiceCouldNotFoundInTheMarkedDocumentsSList") + " : " + ARInvoiceDocEntry + "! ";

                                SAPbobsCOM.GeneralData InvoiceRow = oChildren.Item(i);
                                if (InvoiceRow.GetProperty("U_baseDocT") == "ARInvoice" && InvoiceRow.GetProperty("U_baseDoc") == ARInvoiceDocEntry)
                                {
                                    errorText = null;
                                    return;
                                }
                            }
                        }
                        else
                            errorText = BDOSResources.getTranslate("BaseARInvoiceCouldNotFoundInTheMarkedDocumentsSList") + " : " + ARInvoiceDocEntry + "! ";
                        return;
                    }
                    if (union == false)
                    {
                        //------------------------------------>კორექტირების ფაქტურის მონაცემების შევსება<------------------------------------ 
                        oGeneralData.SetProperty("U_corrInv", "Y");
                        oGeneralData.SetProperty("U_corrType", corrType);

                        //კორექტირების დოკუმენტები                    
                        List<int> connectedDocList = ARInvoice.getAllConnectedDoc(new List<int>() { ARInvoiceDocEntry }, "13", docDate, docDate, docTime, out errorText);
                        int rowCount = connectedDocList.Count();

                        //რეალიზაციის დოკუმენტის თანხა -->
                        double gTotal;
                        double lineVat;
                        ARInvoice.getAmount(ARInvoiceDocEntry, out gTotal, out lineVat, out errorText);

                        decimal amountInvoice = Convert.ToDecimal(gTotal); //რეალიზაციის თანხას უნდა გამოაკლდეს
                        decimal amountTXInvoice = Convert.ToDecimal(lineVat);

                        Dictionary<string, object> taxDocInfo = null;
                        int creditNoteDocEntry = 0;
                        int corrARDocEntry = 0;
                        string corrARInvID = null;
                        string corrARStatus = null;
                        string text;

                        taxDocInfo = getTaxInvoiceSentDocumentInfo(ARInvoiceDocEntry, "ARInvoice", cardCode);
                        if (taxDocInfo != null)
                        {
                            corrARDocEntry = Convert.ToInt32(taxDocInfo["docEntry"]); //კორექტირების ა/ფ-ის Entry
                            corrARInvID = taxDocInfo["invID"].ToString(); //კორექტირების ა/ფ-ის ID
                            corrARStatus = taxDocInfo["status"].ToString(); //კორექტირების ა/ფ-ის სტატუსი
                        }

                        text = BDOSResources.getTranslate("OnTheBaseARInvoice");

                        if (string.IsNullOrEmpty(corrARInvID) && corrARDocEntry == 0)
                        {
                            errorText = BDOSResources.getTranslate("NoTaxInvoiceSaved") + " " + text + "! " + BDOSResources.getTranslate("Document") + " : " + ARInvoiceDocEntry;
                            return;
                        }
                        else if (corrARStatus != "confirmed" && corrARStatus != "correctionConfirmed" && corrARStatus != "primary")
                        {
                            errorText = BDOSResources.getTranslate("NoTaxInvoiceConfirmed") + "! " + BDOSResources.getTranslate("Document") + " : " + corrARDocEntry;
                            return;
                        }

                        if (rowCount == 1) // ეს ნიშნავს რომ რეალიზაციაზე მარტო ერთი сreditNote არის მიბმული. ამიტომ უნდა დავაკორექტიროთ რეალიზაციის ა/ფ.
                        {
                            //თანხები ---------->
                            amount = amountInvoice - amount; //რეალიზაციის თანხას უნდა გამოაკლდეს
                            amountTX = amountTXInvoice - amountTX;
                            //თანხები <----------

                            oGeneralData.SetProperty("U_corrDoc", corrARDocEntry);
                            oGeneralData.SetProperty("U_corrDTxt", corrARDocEntry.ToString());
                            oGeneralData.SetProperty("U_corrDocID", corrARInvID);

                            //საფუძველის ცხრილის შევსება
                            SAPbobsCOM.GeneralDataCollection oChildren = null;
                            oChildren = oGeneralData.Child("BDO_TXS1");

                            SAPbobsCOM.GeneralData oChild = oChildren.Add();

                            oChild.SetProperty("U_baseDocT", "ARCreditNote"); //A/R CreditNote
                            oChild.SetProperty("U_baseDoc", baseDocEntry);
                            oChild.SetProperty("U_baseDTxt", baseDocEntry.ToString());
                            oChild.SetProperty("U_amtBsDc", oRecordSet.Fields.Item("DocTotal").Value); //თანხა დღგ-ის ჩათვლით
                            oChild.SetProperty("U_tAmtBsDc", oRecordSet.Fields.Item("VatSum").Value); //დღგ-ის თანხა
                            oChild.SetProperty("U_wbNumber", wblDocInfo["number"]);
                        }
                        else
                        {
                            int corrDocEntry = 0;
                            string corrInvID = null;
                            string corrStatus = null;
                            text = BDOSResources.getTranslate("OnThePreviousARCorrection");

                            for (int i = 0; i < rowCount; i++)
                            {
                                creditNoteDocEntry = connectedDocList[i];

                                //if (creditNoteDocEntry != baseDocEntry)
                                //{
                                BDO_CNTp = CommonFunctions.getValue("ORIN", "U_BDO_CNTp", "DocEntry", creditNoteDocEntry.ToString()).ToString(); //0=კორექტირება, 1=დაბრუნება

                                taxDocInfo = getTaxInvoiceSentDocumentInfo(creditNoteDocEntry, "ARCreditNote", cardCode);
                                if (taxDocInfo != null)
                                {
                                    corrDocEntry = Convert.ToInt32(taxDocInfo["docEntry"]); //კორექტირების ა/ფ-ის Entry
                                    corrInvID = taxDocInfo["invID"].ToString(); //კორექტირების ა/ფ-ის ID 
                                    corrStatus = taxDocInfo["status"].ToString(); //კორექტირების ა/ფ-ის სტატუსი 
                                }
                                if (creditNoteDocEntry != baseDocEntry && corrDocEntry != 0)
                                {
                                    continue;
                                }

                                if (string.IsNullOrEmpty(corrInvID) && corrDocEntry == 0)
                                {
                                    if (fromDoc)
                                    {
                                        answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("ARCDocumentsOnARIDoYouWantToCreateUnitedTaxInvoiceIncludingTheseDocuments") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), ""); //არსებობს რეალიზაციის დოკუმენტზე რეალიზაციის კორექტირების დოკუმენტები, გსურთ შეიქმნას ერთიანი ფაქტურა ამ დოკუმენტების გათვალისწინებით                   
                                    }
                                    if (answer != 1)
                                    {
                                        errorText = BDOSResources.getTranslate("NoTaxInvoiceSaved") + " " + text + "! " + BDOSResources.getTranslate("Document") + " : " + creditNoteDocEntry;
                                        return;
                                    }
                                }
                                else if (corrStatus != "confirmed" && corrStatus != "correctionConfirmed")
                                {
                                    errorText = BDOSResources.getTranslate("NoTaxInvoiceConfirmed") + "! " + BDOSResources.getTranslate("Document") + " : " + corrDocEntry;
                                    return;
                                }

                                if (BDO_CNTp == "0" && ARInvoiceDocEntry != 0) //კორექტირება
                                {
                                    wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(ARInvoiceDocEntry, "13", out errorText);
                                }
                                else //დაბრუნება
                                {
                                    wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(creditNoteDocEntry, "14", out errorText);
                                }

                                //თანხები ---------->
                                ARCreditNote.getAmount(creditNoteDocEntry, out gTotal, out lineVat, out errorText);
                                amountInvoice = amountInvoice - Convert.ToDecimal(gTotal); //რეალიზაციის თანხას უნდა გამოაკლდეს
                                amountTXInvoice = amountTXInvoice - Convert.ToDecimal(lineVat);
                                //თანხები <----------

                                if (corrDocEntry == 0) //answer == 1
                                {
                                    oGeneralData.SetProperty("U_corrDoc", corrARDocEntry);
                                    oGeneralData.SetProperty("U_corrDTxt", corrARDocEntry.ToString());
                                    oGeneralData.SetProperty("U_corrDocID", corrARInvID);
                                }
                                else
                                {
                                    oGeneralData.SetProperty("U_corrDoc", corrDocEntry);
                                    oGeneralData.SetProperty("U_corrDTxt", corrDocEntry.ToString());
                                    oGeneralData.SetProperty("U_corrDocID", corrInvID);
                                }

                                //საფუძველის ცხრილის შევსება
                                SAPbobsCOM.GeneralDataCollection oChildren = null;
                                oChildren = oGeneralData.Child("BDO_TXS1");

                                SAPbobsCOM.GeneralData oChild = oChildren.Add();

                                oChild.SetProperty("U_baseDocT", "ARCreditNote"); //A/R CreditNote
                                oChild.SetProperty("U_baseDoc", creditNoteDocEntry);
                                oChild.SetProperty("U_baseDTxt", creditNoteDocEntry.ToString());
                                oChild.SetProperty("U_amtBsDc", gTotal); //თანხა დღგ-ის ჩათვლით
                                oChild.SetProperty("U_tAmtBsDc", lineVat); //დღგ-ის თანხა
                                oChild.SetProperty("U_wbNumber", wblDocInfo["number"]);
                                //}
                            }
                        }
                        //------------------------------------>კორექტირების ფაქტურის მონაცემების შევსება<------------------------------------                   

                        //საბოლოო თანხის შევსება
                        oGeneralData.SetProperty("U_amount", Convert.ToDouble(amount)); //თანხა დღგ-ის ჩათვლით
                        oGeneralData.SetProperty("U_amountTX", Convert.ToDouble(amountTX)); //დღგ-ის თანხა
                        oGeneralData.SetProperty("U_amtOutTX", Convert.ToDouble(amount - amountTX)); //თანხა დღგ-ის გარეშე

                        try
                        {
                            var response = oGeneralService.Add(oGeneralData);
                            var docEntry = response.GetProperty("DocEntry");
                            newDocEntry = Convert.ToInt32(docEntry);
                        }
                        catch (Exception ex)
                        {
                            errorText = BDOSResources.getTranslate("ErrorDocumentAddEdit") + " : " + ex.Message;
                            return;
                        }
                    }
                    oRecordSet.MoveNext();
                    break;
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

        private static void createDocumentInvoiceCorrectionInvoiceType(int baseDocEntry, string corrType, bool fromDoc,
            int answer, SAPbobsCOM.GeneralData oGeneralData, bool union, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;

            SAPbobsCOM.Recordset oRecordSet = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "SELECT " +
                           "\"OCSI\".\"DocEntry\", " +
                           "\"OCSI\".\"DocDate\", " +
                           "\"OCSI\".\"DocTime\", " +
                           "\"OCSI\".\"CardCode\", " +
                           "\"OCSI\".\"U_BDOSCITp\", " +
                           "\"OCRD\".\"CardName\", " +
                           "\"OCRD\".\"LicTradNum\", " +
                           "SUM(\"CSI1\".\"GTotal\") as \"DocTotal\", " +
                           "SUM(\"CSI1\".\"LineVat\") as \"VatSum\"  " +
                           "FROM \"CSI1\"  " +
                           "left join \"OCSI\" on \"CSI1\".\"DocEntry\" = \"OCSI\".\"DocEntry\"  " +
                           "left join \"OCRD\" on \"OCSI\".\"CardCode\" = \"OCRD\".\"CardCode\"  " +
                           "WHERE \"CSI1\".\"DocEntry\" = '" + baseDocEntry + "'  " +
                           "GROUP BY \"OCSI\".\"DocEntry\", " +
                           "\"OCSI\".\"DocDate\", " +
                           "\"OCSI\".\"DocTime\", " +
                           "\"OCSI\".\"CardCode\", " +
                           "\"OCSI\".\"U_BDOSCITp\", " +
                           "\"OCRD\".\"CardName\", " +
                           "\"OCRD\".\"LicTradNum\"  ";

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    decimal amount = Convert.ToDecimal(oRecordSet.Fields.Item("DocTotal").Value);
                    decimal amountTX = Convert.ToDecimal(oRecordSet.Fields.Item("VatSum").Value);
                    decimal amtOutTX = amount - amountTX;

                    DateTime docDate = oRecordSet.Fields.Item("DocDate").Value;
                    int docTime = oRecordSet.Fields.Item("DocTime").Value;
                    string cardCode = oRecordSet.Fields.Item("CardCode").Value.ToString();
                    string BDOSCITp = oRecordSet.Fields.Item("U_BDOSCITp").Value.ToString();

                    SAPbobsCOM.CompanyService oCompanyService = null;
                    SAPbobsCOM.GeneralService oGeneralService = null;

                    if (union == false)
                    {
                        oCompanyService = Program.oCompany.GetCompanyService();
                        oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXS_D");
                        oGeneralData =
                            ((SAPbobsCOM.GeneralData) (oGeneralService.GetDataInterface(SAPbobsCOM
                                .GeneralServiceDataInterfaces.gsGeneralData)));

                        oGeneralData.SetProperty("U_opDate", new DateTime(docDate.Year, docDate.Month, 1));
                        oGeneralData.SetProperty("U_cardCode", cardCode);
                        oGeneralData.SetProperty("U_cardCodeN", oRecordSet.Fields.Item("CardName").Value.ToString());
                        oGeneralData.SetProperty("U_cardCodeT", oRecordSet.Fields.Item("LicTradNum").Value.ToString());
                        oGeneralData.SetProperty("U_status", "empty");
                    }

                    ArCorrectionInvoice.GetBaseDoc(baseDocEntry, out int ARInvoiceDocEntry);

                    Dictionary<string, string> wblDocInfo = null;
                    if (BDOSCITp == "0" && ARInvoiceDocEntry != 0) //კორექტირება                   
                    {
                        wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(ARInvoiceDocEntry, "13", out errorText);
                    }

                    else
                    {
                        wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(baseDocEntry, "165", out errorText);
                    }

                    if (ARInvoiceDocEntry == 0)
                    {
                        if (union == false)
                        {
                            errorText = BDOSResources.getTranslate("BaseARInvoiceCouldNotFound") + "! ";
                            return;
                        }

                        if (fromDoc)
                        {
                            answer = Program.uiApp.MessageBox(
                                BDOSResources.getTranslate("BaseARInvoiceCouldNotFound") + "! " +
                                BDOSResources.getTranslate("ContinueCreatingTaxInvoice") + "?", 1,
                                BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                        }

                        if (answer == 1)
                        {
                            try
                            {
                                //საფუძველის ცხრილის შევსება
                                SAPbobsCOM.GeneralDataCollection oChildren = null;
                                oChildren = oGeneralData.Child("BDO_TXS1");

                                SAPbobsCOM.GeneralData oChild = oChildren.Add();

                                oChild.SetProperty("U_baseDocT", "ARCorrectionInvoice"); //A/R Correction Invoice
                                oChild.SetProperty("U_baseDoc", baseDocEntry);
                                oChild.SetProperty("U_baseDTxt", baseDocEntry.ToString());
                                oChild.SetProperty("U_amtBsDc",
                                    oRecordSet.Fields.Item("DocTotal").Value); //თანხა დღგ-ის ჩათვლით
                                oChild.SetProperty("U_tAmtBsDc", oRecordSet.Fields.Item("VatSum").Value); //დღგ-ის თანხა
                                oChild.SetProperty("U_wbNumber", wblDocInfo["number"]);

                                if (union == false)
                                {
                                    oGeneralData.SetProperty("U_amount",
                                        Convert.ToDouble(amount)); //თანხა დღგ-ის ჩათვლით
                                    oGeneralData.SetProperty("U_amountTX", Convert.ToDouble(amountTX)); //დღგ-ის თანხა
                                    oGeneralData.SetProperty("U_amtOutTX",
                                        Convert.ToDouble(amount - amountTX)); //თანხა დღგ-ის გარეშე

                                    var response = oGeneralService.Add(oGeneralData);
                                    var docEntry = response.GetProperty("DocEntry");
                                    newDocEntry = Convert.ToInt32(docEntry);
                                }
                            }
                            catch (Exception ex)
                            {
                                errorText = BDOSResources.getTranslate("ErrorDocumentAddEdit") + " : " + ex.Message;
                                return;
                            }
                        }

                        return;
                    }
                    else if (union)
                    {
                        SAPbobsCOM.GeneralDataCollection oChildren = null;
                        oChildren = oGeneralData.Child("BDO_TXS1");
                        int count = oChildren.Count;
                        if (count > 0)
                        {
                            for (int i = 0; i < count; i++)
                            {
                                errorText =
                                    BDOSResources.getTranslate("BaseARInvoiceCouldNotFoundInTheMarkedDocumentsSList") +
                                    " : " + ARInvoiceDocEntry + "! ";

                                SAPbobsCOM.GeneralData InvoiceRow = oChildren.Item(i);
                                if (InvoiceRow.GetProperty("U_baseDocT") == "ARInvoice" &&
                                    InvoiceRow.GetProperty("U_baseDoc") == ARInvoiceDocEntry)
                                {
                                    errorText = null;
                                    return;
                                }
                            }
                        }
                        else
                            errorText =
                                BDOSResources.getTranslate("BaseARInvoiceCouldNotFoundInTheMarkedDocumentsSList") +
                                " : " + ARInvoiceDocEntry + "! ";

                        return;
                    }

                    if (union == false)
                    {
                        //------------------------------------>კორექტირების ფაქტურის მონაცემების შევსება<------------------------------------ 
                        oGeneralData.SetProperty("U_corrInv", "Y");
                        oGeneralData.SetProperty("U_corrType", corrType);

                        //კორექტირების დოკუმენტები                    
                        List<int> connectedDocList = ARInvoice.getAllConnectedARCorrectionDoc(new List<int>() {ARInvoiceDocEntry},
                            "13", docDate, docDate, docTime, out errorText);
                        int rowCount = connectedDocList.Count();

                        //რეალიზაციის დოკუმენტის თანხა -->
                        double gTotal;
                        double lineVat;
                        ARInvoice.getAmount(ARInvoiceDocEntry, out gTotal, out lineVat, out errorText);

                        decimal amountInvoice = Convert.ToDecimal(gTotal); //რეალიზაციის თანხას უნდა გამოაკლდეს
                        decimal amountTXInvoice = Convert.ToDecimal(lineVat);

                        Dictionary<string, object> taxDocInfo = null;
                        int correctionInvoiceDocEntry = 0;
                        int corrARDocEntry = 0;
                        string corrARInvID = null;
                        string corrARStatus = null;
                        string text;

                        taxDocInfo = getTaxInvoiceSentDocumentInfo(ARInvoiceDocEntry, "ARInvoice", cardCode);
                        if (taxDocInfo != null)
                        {
                            corrARDocEntry = Convert.ToInt32(taxDocInfo["docEntry"]); //კორექტირების ა/ფ-ის Entry
                            corrARInvID = taxDocInfo["invID"].ToString(); //კორექტირების ა/ფ-ის ID
                            corrARStatus = taxDocInfo["status"].ToString(); //კორექტირების ა/ფ-ის სტატუსი
                        }

                        text = BDOSResources.getTranslate("OnTheBaseARInvoice");

                        if (string.IsNullOrEmpty(corrARInvID) && corrARDocEntry == 0)
                        {
                            errorText = BDOSResources.getTranslate("NoTaxInvoiceSaved") + " " + text + "! " +
                                        BDOSResources.getTranslate("Document") + " : " + ARInvoiceDocEntry;
                            return;
                        }
                        else if (corrARStatus != "confirmed" && corrARStatus != "correctionConfirmed" &&
                                 corrARStatus != "primary")
                        {
                            errorText = BDOSResources.getTranslate("NoTaxInvoiceConfirmed") + "! " +
                                        BDOSResources.getTranslate("Document") + " : " + corrARDocEntry;
                            return;
                        }

                        if (rowCount == 1
                        ) // ეს ნიშნავს რომ რეალიზაციაზე მარტო ერთი Correction Invoice არის მიბმული. ამიტომ უნდა დავაკორექტიროთ რეალიზაციის ა/ფ.
                        {
                            //თანხები ---------->
                            amount = amountInvoice - amount; //რეალიზაციის თანხას უნდა გამოაკლდეს
                            amountTX = amountTXInvoice - amountTX;
                            //თანხები <----------

                            oGeneralData.SetProperty("U_corrDoc", corrARDocEntry);
                            oGeneralData.SetProperty("U_corrDTxt", corrARDocEntry.ToString());
                            oGeneralData.SetProperty("U_corrDocID", corrARInvID);

                            //საფუძველის ცხრილის შევსება
                            SAPbobsCOM.GeneralDataCollection oChildren = null;
                            oChildren = oGeneralData.Child("BDO_TXS1");

                            SAPbobsCOM.GeneralData oChild = oChildren.Add();

                            oChild.SetProperty("U_baseDocT", "ARCorrectionInvoice"); //A/R Correction Invoice
                            oChild.SetProperty("U_baseDoc", baseDocEntry);
                            oChild.SetProperty("U_baseDTxt", baseDocEntry.ToString());
                            oChild.SetProperty("U_amtBsDc",
                                oRecordSet.Fields.Item("DocTotal").Value); //თანხა დღგ-ის ჩათვლით
                            oChild.SetProperty("U_tAmtBsDc", oRecordSet.Fields.Item("VatSum").Value); //დღგ-ის თანხა
                            oChild.SetProperty("U_wbNumber", wblDocInfo["number"]);
                        }
                        else
                        {
                            int corrDocEntry = 0;
                            string corrInvID = null;
                            string corrStatus = null;
                            text = BDOSResources.getTranslate("OnThePreviousARCorrection");

                            for (int i = 0; i < rowCount; i++)
                            {
                                correctionInvoiceDocEntry = connectedDocList[i];

                                BDOSCITp = CommonFunctions.getValue("OCSI", "U_BDOSCITp", "DocEntry",
                                    correctionInvoiceDocEntry.ToString()).ToString(); //0=კორექტირება, 1=დაბრუნება

                                taxDocInfo = getTaxInvoiceSentDocumentInfo(correctionInvoiceDocEntry,
                                    "ARCorrectionInvoice", cardCode);
                                if (taxDocInfo != null)
                                {
                                    corrDocEntry = Convert.ToInt32(taxDocInfo["docEntry"]); //კორექტირების ა/ფ-ის Entry
                                    corrInvID = taxDocInfo["invID"].ToString(); //კორექტირების ა/ფ-ის ID 
                                    corrStatus = taxDocInfo["status"].ToString(); //კორექტირების ა/ფ-ის სტატუსი 
                                }

                                if (correctionInvoiceDocEntry != baseDocEntry && corrDocEntry != 0)
                                {
                                    continue;
                                }

                                if (string.IsNullOrEmpty(corrInvID) && corrDocEntry == 0)
                                {
                                    if (fromDoc)
                                    {
                                        answer = Program.uiApp.MessageBox(
                                            BDOSResources.getTranslate(
                                                "ARCDocumentsOnARIDoYouWantToCreateUnitedTaxInvoiceIncludingTheseDocuments") +
                                            "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"),
                                            ""); //არსებობს რეალიზაციის დოკუმენტზე რეალიზაციის კორექტირების დოკუმენტები, გსურთ შეიქმნას ერთიანი ფაქტურა ამ დოკუმენტების გათვალისწინებით                   
                                    }

                                    if (answer != 1)
                                    {
                                        errorText = BDOSResources.getTranslate("NoTaxInvoiceSaved") + " " + text +
                                                    "! " + BDOSResources.getTranslate("Document") + " : " +
                                                    correctionInvoiceDocEntry;
                                        return;
                                    }
                                }
                                else if (corrStatus != "confirmed" && corrStatus != "correctionConfirmed")
                                {
                                    errorText = BDOSResources.getTranslate("NoTaxInvoiceConfirmed") + "! " +
                                                BDOSResources.getTranslate("Document") + " : " + corrDocEntry;
                                    return;
                                }

                                if (BDOSCITp == "0" && ARInvoiceDocEntry != 0) //კორექტირება
                                {
                                    wblDocInfo =
                                        BDO_Waybills.getWaybillDocumentInfo(ARInvoiceDocEntry, "13", out errorText);
                                }
                                else //დაბრუნება
                                {
                                    wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(correctionInvoiceDocEntry, "165",
                                        out errorText);
                                }

                                //თანხები ---------->
                                ArCorrectionInvoice.GetAmount(correctionInvoiceDocEntry, out gTotal, out lineVat,
                                    out errorText);
                                amountInvoice =
                                    amountInvoice - Convert.ToDecimal(gTotal); //რეალიზაციის თანხას უნდა გამოაკლდეს
                                amountTXInvoice = amountTXInvoice - Convert.ToDecimal(lineVat);
                                //თანხები <----------

                                if (corrDocEntry == 0) //answer == 1
                                {
                                    oGeneralData.SetProperty("U_corrDoc", corrARDocEntry);
                                    oGeneralData.SetProperty("U_corrDTxt", corrARDocEntry.ToString());
                                    oGeneralData.SetProperty("U_corrDocID", corrARInvID);
                                }
                                else
                                {
                                    oGeneralData.SetProperty("U_corrDoc", corrDocEntry);
                                    oGeneralData.SetProperty("U_corrDTxt", corrDocEntry.ToString());
                                    oGeneralData.SetProperty("U_corrDocID", corrInvID);
                                }

                                //საფუძველის ცხრილის შევსება
                                SAPbobsCOM.GeneralDataCollection oChildren = null;
                                oChildren = oGeneralData.Child("BDO_TXS1");

                                SAPbobsCOM.GeneralData oChild = oChildren.Add();

                                oChild.SetProperty("U_baseDocT", "ARCorrectionInvoice"); //A/R Correction Invoice
                                oChild.SetProperty("U_baseDoc", correctionInvoiceDocEntry);
                                oChild.SetProperty("U_baseDTxt", correctionInvoiceDocEntry.ToString());
                                oChild.SetProperty("U_amtBsDc", gTotal); //თანხა დღგ-ის ჩათვლით
                                oChild.SetProperty("U_tAmtBsDc", lineVat); //დღგ-ის თანხა
                                oChild.SetProperty("U_wbNumber", wblDocInfo["number"]);
                                //}
                            }
                        }
                        //------------------------------------>კორექტირების ფაქტურის მონაცემების შევსება<------------------------------------                   

                        //საბოლოო თანხის შევსება
                        oGeneralData.SetProperty("U_amount", Convert.ToDouble(amount)); //თანხა დღგ-ის ჩათვლით
                        oGeneralData.SetProperty("U_amountTX", Convert.ToDouble(amountTX)); //დღგ-ის თანხა
                        oGeneralData.SetProperty("U_amtOutTX",
                            Convert.ToDouble(amount - amountTX)); //თანხა დღგ-ის გარეშე

                        try
                        {
                            var response = oGeneralService.Add(oGeneralData);
                            var docEntry = response.GetProperty("DocEntry");
                            newDocEntry = Convert.ToInt32(docEntry);
                        }
                        catch (Exception ex)
                        {
                            errorText = BDOSResources.getTranslate("ErrorDocumentAddEdit") + " : " + ex.Message;
                            return;
                        }
                    }

                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        private static void createDocumentInvoiceARDownPaymentRequestType(int baseDocEntry, string corrType, bool fromDoc, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;

            bool primary;
            DataTable confirmedInvoices;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT
            	 ""ODPI"".""DocDate"",
            	 ""ODPI"".""DocEntry"",
            	 ""ODPI"".""DocNum"",
            	 ""ODPI"".""Posted"",
            	 ""ODPI"".""CEECFlag"",
            	 ""ODPI"".""DpmStatus"",
            	 SUM(""DPI1"".""U_BDOSDPMAmt"") AS ""DocTotal"",
            	 SUM(""DPI1"".""U_BDOSDPMVat"") AS ""VatSum"",
            	 ""OCRD"".""CardName"",
                 ""OCRD"".""CardCode"",
            	 ""OCRD"".""LicTradNum"" 
            FROM ""ODPI"" AS ""ODPI"" 
            INNER JOIN ""DPI1"" AS ""DPI1"" ON ""ODPI"".""DocEntry"" = ""DPI1"".""DocEntry"" 
            LEFT JOIN ""OCRD"" AS ""OCRD"" ON ""ODPI"".""CardCode"" = ""OCRD"".""CardCode"" 
            WHERE ""ODPI"".""DocEntry"" = " + baseDocEntry + " " +
            @"--AND ""ODPI"".""Posted"" = 'Y' 
            GROUP BY ""ODPI"".""DocDate"",
            	 ""ODPI"".""DocEntry"",
            	 ""ODPI"".""DocNum"",
            	 ""ODPI"".""Posted"",
            	 ""ODPI"".""CEECFlag"",
            	 ""ODPI"".""DpmStatus"",
            	 ""OCRD"".""CardName"",
                 ""OCRD"".""CardCode"",
            	 ""OCRD"".""LicTradNum""";
            try
            {
                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    SAPbobsCOM.CompanyService oCompanyService = null;
                    SAPbobsCOM.GeneralService oGeneralService = null;
                    SAPbobsCOM.GeneralData oGeneralData = null;
                    SAPbobsCOM.GeneralDataParams oGeneralParams = null;

                    oCompanyService = Program.oCompany.GetCompanyService();
                    oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXS_D");
                    oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                    if (newDocEntry != 0)
                    {
                        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        oGeneralParams.SetProperty("DocEntry", newDocEntry);
                        oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                    }

                    DateTime docDate = oRecordSet.Fields.Item("DocDate").Value;

                    if (ARDownPaymentRequest.checkDocumentForTaxInvoice(baseDocEntry, docDate, docDate, out primary, out confirmedInvoices, out errorText) == false)
                    {
                        return;
                    }

                    oGeneralData.SetProperty("U_opDate", new DateTime(docDate.Year, docDate.Month, 1));
                    string cardCode = oRecordSet.Fields.Item("CardCode").Value.ToString();
                    oGeneralData.SetProperty("U_cardCode", cardCode);
                    oGeneralData.SetProperty("U_cardCodeN", oRecordSet.Fields.Item("CardName").Value.ToString());
                    oGeneralData.SetProperty("U_cardCodeT", oRecordSet.Fields.Item("LicTradNum").Value.ToString());
                    oGeneralData.SetProperty("U_status", "empty");
                    oGeneralData.SetProperty("U_downPaymnt", "Y");

                    //------------------------------------>კორექტირების ფაქტურის მონაცემების შევსება<------------------------------------
                    if (primary == false)
                    {
                        DataRow taxDataRow = confirmedInvoices.Rows[0];
                        oGeneralData.SetProperty("U_corrDoc", taxDataRow["InvDocEntry"]);
                        oGeneralData.SetProperty("U_corrDTxt", taxDataRow["InvDocEntry"].ToString());
                        oGeneralData.SetProperty("U_corrDocID", taxDataRow["U_invID"].ToString());
                        oGeneralData.SetProperty("U_corrInv", "Y");
                        if (string.IsNullOrEmpty(corrType) == false)
                        {
                            oGeneralData.SetProperty("U_corrType", corrType);
                        }
                    }
                    //------------------------------------>კორექტირების ფაქტურის მონაცემების შევსება<------------------------------------

                    decimal amount = Convert.ToDecimal(oRecordSet.Fields.Item("DocTotal").Value);
                    decimal amountTX = Convert.ToDecimal(oRecordSet.Fields.Item("VatSum").Value);
                    decimal amtOutTX = amount - amountTX;

                    SAPbobsCOM.GeneralDataCollection oChildren = null;

                    oChildren = oGeneralData.Child("BDO_TXS1");

                    SAPbobsCOM.GeneralData oChild = oChildren.Add();

                    oChild.SetProperty("U_baseDocT", "ARDownPaymentRequest");
                    oChild.SetProperty("U_baseDoc", oRecordSet.Fields.Item("DocEntry").Value);
                    oChild.SetProperty("U_baseDTxt", oRecordSet.Fields.Item("DocEntry").Value.ToString());
                    oChild.SetProperty("U_amtBsDc", oRecordSet.Fields.Item("DocTotal").Value); //თანხა დღგ-ის ჩათვლით
                    oChild.SetProperty("U_tAmtBsDc", oRecordSet.Fields.Item("VatSum").Value); //დღგ-ის თანხა
                    oGeneralData.SetProperty("U_amount", Convert.ToDouble(amount)); //თანხა დღგ-ის ჩათვლით
                    oGeneralData.SetProperty("U_amountTX", Convert.ToDouble(amountTX)); //დღგ-ის თანხა
                    oGeneralData.SetProperty("U_amtOutTX", Convert.ToDouble(amtOutTX)); //თანხა დღგ-ის გარეშე

                    try
                    {
                        if (newDocEntry != 0)
                        {
                            oGeneralService.Update(oGeneralData);
                        }
                        else
                        {
                            var response = oGeneralService.Add(oGeneralData);
                            var docEntry = response.GetProperty("DocEntry");
                            newDocEntry = Convert.ToInt32(docEntry);
                        }
                    }
                    catch (Exception ex)
                    {
                        errorText = BDOSResources.getTranslate("ErrorDocumentAddEdit") + " : " + ex.Message;
                    }

                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
        }

        private static void createDocumentInvoiceARDownPaymentVAT(int baseDocEntry, string corrType, bool fromDoc, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT
            	 ""@BDOSARDV"".""U_DocDate"",
            	 ""@BDOSARDV"".""DocEntry"",
            	 ""@BDOSARDV"".""DocNum"",
            	 
            	 SUM(""@BDOSARDV"".""U_GrsAmnt"") AS ""DocTotal"",
            	 SUM(""@BDOSARDV"".""U_VatAmount"") AS ""VatSum"",
            	 ""OCRD"".""CardName"",
                 ""OCRD"".""CardCode"",
            	 ""OCRD"".""LicTradNum"" 
            FROM ""@BDOSARDV"" AS ""@BDOSARDV"" 
            LEFT JOIN ""OCRD"" AS ""OCRD"" ON ""@BDOSARDV"".""U_cardCode"" = ""OCRD"".""CardCode"" 
            WHERE ""@BDOSARDV"".""DocEntry"" = " + baseDocEntry + " " +
            @"--AND ""@BDOSARDV"".""Posted"" = 'Y' 
            GROUP BY ""@BDOSARDV"".""U_DocDate"",
            	 ""@BDOSARDV"".""DocEntry"",
            	 ""@BDOSARDV"".""DocNum"",
            	 
            	 ""OCRD"".""CardName"",
                 ""OCRD"".""CardCode"",
            	 ""OCRD"".""LicTradNum""";
            try
            {
                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    SAPbobsCOM.CompanyService oCompanyService = null;
                    SAPbobsCOM.GeneralService oGeneralService = null;
                    SAPbobsCOM.GeneralData oGeneralData = null;
                    SAPbobsCOM.GeneralDataParams oGeneralParams = null;

                    oCompanyService = Program.oCompany.GetCompanyService();
                    oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXS_D");
                    oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                    if (newDocEntry != 0)
                    {
                        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        oGeneralParams.SetProperty("DocEntry", newDocEntry);
                        oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                    }

                    DateTime docDate = oRecordSet.Fields.Item("U_DocDate").Value;

                    oGeneralData.SetProperty("U_opDate", new DateTime(docDate.Year, docDate.Month, 1));
                    string cardCode = oRecordSet.Fields.Item("CardCode").Value.ToString();
                    oGeneralData.SetProperty("U_cardCode", cardCode);
                    oGeneralData.SetProperty("U_cardCodeN", oRecordSet.Fields.Item("CardName").Value.ToString());
                    oGeneralData.SetProperty("U_cardCodeT", oRecordSet.Fields.Item("LicTradNum").Value.ToString());
                    oGeneralData.SetProperty("U_status", "empty");
                    oGeneralData.SetProperty("U_downPaymnt", "Y");

                    decimal amount = Convert.ToDecimal(oRecordSet.Fields.Item("DocTotal").Value);
                    decimal amountTX = Convert.ToDecimal(oRecordSet.Fields.Item("VatSum").Value);
                    decimal amtOutTX = amount - amountTX;

                    SAPbobsCOM.GeneralDataCollection oChildren = null;

                    oChildren = oGeneralData.Child("BDO_TXS1");

                    SAPbobsCOM.GeneralData oChild = oChildren.Add();

                    oChild.SetProperty("U_baseDocT", "ARDownPaymentVAT");
                    oChild.SetProperty("U_baseDoc", oRecordSet.Fields.Item("DocEntry").Value);
                    oChild.SetProperty("U_baseDTxt", oRecordSet.Fields.Item("DocEntry").Value.ToString());
                    oChild.SetProperty("U_amtBsDc", oRecordSet.Fields.Item("DocTotal").Value); //თანხა დღგ-ის ჩათვლით
                    oChild.SetProperty("U_tAmtBsDc", oRecordSet.Fields.Item("VatSum").Value); //დღგ-ის თანხა
                    oGeneralData.SetProperty("U_amount", Convert.ToDouble(amount)); //თანხა დღგ-ის ჩათვლით
                    oGeneralData.SetProperty("U_amountTX", Convert.ToDouble(amountTX)); //დღგ-ის თანხა
                    oGeneralData.SetProperty("U_amtOutTX", Convert.ToDouble(amtOutTX)); //თანხა დღგ-ის გარეშე

                    try
                    {
                        if (newDocEntry != 0)
                        {
                            oGeneralService.Update(oGeneralData);
                        }
                        else
                        {
                            var response = oGeneralService.Add(oGeneralData);
                            var docEntry = response.GetProperty("DocEntry");
                            newDocEntry = Convert.ToInt32(docEntry);
                        }
                    }
                    catch (Exception ex)
                    {
                        errorText = BDOSResources.getTranslate("ErrorDocumentAddEdit") + " : " + ex.Message;
                    }

                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
        }

        /// <summary>აბრუნებს რეალიზაციაზე გამოწერლ ყველა сreditNote-ს დასორტირებულს თარიღით და DocEntry</summary>
        /// <param name="Program.oCompany"></param>
        /// <param name="docEntry">сreditNote-ის docEntry</param>
        /// <param name="docDate">сreditNote-ის docDate</param>
        /// <param name="ARInvoiceDocEntry">საფუძველი ARInvoice-ის docEntry</param>
        /// <param name="cardCode">сreditNote-ის cardCode</param>
        /// <param name="сreditNotes"></param>
        /// <param name="errorText"></param>
        public static DataTable getCorrDocs(int docEntry, DateTime docDate, DateTime docDateForMonth, int ARInvoiceDocEntry, string cardCode, out string errorText)
        {
            errorText = null;
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            DataTable сreditNotes = null;

            DateTime firstDay = new DateTime(docDateForMonth.Year, docDateForMonth.Month, 1);
            DateTime lastDay = firstDay.AddMonths(1).AddDays(-1);

            string queryCorrDoc = "SELECT * FROM " +
                    "(SELECT DISTINCT " +
                         "\"ORIN\".\"DocEntry\" AS \"DocEntry\", " +
                         "\"ORIN\".\"DocDate\" AS \"DocDate\", " +
                         "\"ORIN\".\"BaseEntry\" AS \"BaseEntry\", " +
                         "\"ORIN\".\"U_BDO_CNTp\" AS \"U_BDO_CNTp\", " +
                         "\"BDO_TAXS\".\"DocEntry\" AS \"CorrDocEntry\", " +
                         "\"BDO_TAXS\".\"U_invID\" AS \"U_invID\", " +
                         "\"BDO_TAXS\".\"U_status\" AS \"U_status\" " +
                       "FROM ( " +
                        "(SELECT " +
                         "\"RIN1\".\"DocEntry\" AS \"DocEntry\", " +
                         "\"RIN1\".\"DocDate\" AS \"DocDate\", " +
                         "\"RIN1\".\"BaseEntry\" AS \"BaseEntry\", " +
                         "\"ORIN\".\"U_BDO_CNTp\" AS \"U_BDO_CNTp\" " +
                         "FROM \"ORIN\" " +
                         "INNER JOIN \"RIN1\" AS \"RIN1\" " +
                         "ON \"ORIN\".\"DocEntry\" = \"RIN1\".\"DocEntry\" " +
                         "WHERE \"RIN1\".\"BaseEntry\" = '" + ARInvoiceDocEntry + "' AND \"RIN1\".\"BaseType\" = '13' AND \"ORIN\".\"CardCode\" = N'" + cardCode + "' " + "  AND (\"CANCELED\"='N') " +
                         ((docDate == new DateTime() || docEntry == 0) ? "" : "AND \"ORIN\".\"DocDate\" <= '" + docDate.ToString("yyyyMMdd") + "' " +
                         "AND \"ORIN\".\"DocDate\" >= '" + firstDay.ToString("yyyyMMdd") + "' AND \"ORIN\".\"DocDate\" <= '" + lastDay.ToString("yyyyMMdd") + "' " +
                         "AND \"ORIN\".\"DocEntry\" <= '" + docEntry + "'") + ") AS \"ORIN\" " +
                         "LEFT JOIN " +
                         "(SELECT " +
                         "\"BDO_TAXS\".\"U_invID\" AS \"U_invID\", " +
                         "\"BDO_TAXS\".\"DocEntry\" AS \"DocEntry\", " +
                         "\"BDO_TXS1\".\"U_baseDoc\" AS \"U_baseDoc\", " +
                         "\"BDO_TXS1\".\"U_baseDocT\" AS \"U_baseDocT\", " +
                         "\"BDO_TAXS\".\"U_status\" AS \"U_status\" " +
                         "FROM \"@BDO_TAXS\" AS \"BDO_TAXS\" " +
                         "INNER JOIN \"@BDO_TXS1\" AS \"BDO_TXS1\" " +
                         "ON \"BDO_TAXS\".\"DocEntry\" = \"BDO_TXS1\".\"DocEntry\" " +
                         "WHERE (\"BDO_TXS1\".\"U_baseDocT\" = 'ARCreditNote') AND \"BDO_TAXS\".\"U_status\" NOT IN ('removed', 'canceled', 'denied', 'paper') AND \"BDO_TAXS\".\"U_cardCode\" = N'" + cardCode + "') AS \"BDO_TAXS\" " +

                       "ON \"ORIN\".\"DocEntry\" = \"BDO_TAXS\".\"U_baseDoc\")) " +

                    "ORDER BY \"DocDate\", \"DocEntry\" DESC";

            try
            {
                oRecordSet.DoQuery(queryCorrDoc);

                int recordCount = oRecordSet.RecordCount;
                if (recordCount > 0)
                {
                    сreditNotes = new DataTable(); //ამაში დაგროვდება ის CreditNote - ები რომლებიც არის რეალიზაციაზე გაფორმებული.
                    сreditNotes.Columns.Add("creditNote", typeof(SAPbobsCOM.Documents)); //creditNote ის დოკუმენტი
                    сreditNotes.Columns.Add("baseEntry", typeof(int)); //საფუძველი დოკუმენტის Entry(A/R Invoice) იგივე ARInvoiceDocEntry
                    сreditNotes.Columns.Add("corrDocEntry", typeof(int)); //კორექტირების ა/ფ-ის Entry (creditNote-ზე გამოწერილი)
                    сreditNotes.Columns.Add("invID", typeof(string)); //კორექტირების ა/ფ-ის ID (creditNote-ზე გამოწერილი)
                    сreditNotes.Columns.Add("status", typeof(string)); //კორექტირების ა/ფ-ის სტატუსი (creditNote-ზე გამოწერილი)
                    сreditNotes.Columns.Add("BDO_CNTp", typeof(string));

                    while (!oRecordSet.EoF)
                    {
                        SAPbobsCOM.Documents oCreditNote;
                        oCreditNote = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                        oCreditNote.GetByKey(oRecordSet.Fields.Item("DocEntry").Value);

                        DataRow taxDataRow = сreditNotes.Rows.Add();
                        taxDataRow["creditNote"] = oCreditNote;
                        taxDataRow["baseEntry"] = oRecordSet.Fields.Item("BaseEntry").Value;
                        taxDataRow["corrDocEntry"] = oRecordSet.Fields.Item("CorrDocEntry").Value;
                        taxDataRow["invID"] = oRecordSet.Fields.Item("U_invID").Value;
                        taxDataRow["status"] = oRecordSet.Fields.Item("U_status").Value;
                        taxDataRow["BDO_CNTp"] = oRecordSet.Fields.Item("U_BDO_CNTp").Value;

                        oRecordSet.MoveNext();
                    }
                }
                return сreditNotes;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction)
            {
                return;
            }

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_TAXS_D")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
                {
                    formDataLoad(oForm, out errorText);
                    setVisibleFormItems(oForm, out errorText);
                }

                if ((BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) & BusinessObjectInfo.BeforeAction)
                {
                    if (Program.cancellationTrans && Program.canceledDocEntry != 0)
                    {
                        Program.canceledDocEntry = 0;
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
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_VISIBLE)
                    {
                        setSizeForm(oForm);
                        oForm.Title = BDOSResources.getTranslate("TaxInvoiceSent");
                        Program.FORM_LOAD_FOR_VISIBLE = false;
                    }
                    setVisibleFormItems(oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    resizeForm(oForm);
                }

                if ((pVal.ItemUID == "cardCodeE" || pVal.ItemUID == "wblMTR") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, pVal.Row, out errorText);
                }

                if (pVal.ItemUID == "wblMTR" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
                {
                    matrixColumnSetCfl(oForm, pVal, out errorText);
                }

                if (pVal.ItemUID == "wblMTR" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                {
                    matrixColumnSetCfl(oForm, pVal, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    itemPressed(oForm, pVal, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.ItemUID == "addMTRB" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                {
                    addMatrixRow(oForm);
                }

                if (pVal.ItemUID == "delMTRB" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                {
                    deleteMatrixRow(oForm);
                    CalculateAmount(oForm);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                {
                    oForm.Freeze(true);
                    comboSelect(oForm, pVal, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE & pVal.BeforeAction == false)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE)
                    {
                        oForm.Freeze(true);
                        setVisibleFormItems(oForm, out errorText);
                        formDataLoad(oForm, out errorText);
                        oForm.Freeze(false);
                        //oForm.Update();
                        Program.FORM_LOAD_FOR_ACTIVATE = false;
                    }
                }

                if (pVal.ItemUID == "operationB" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction)
                {
                    oForm.Freeze(true);
                    setValidValuesBtnCombo(oForm, out errorText);
                    oForm.Freeze(false);
                    //oForm.Update();
                }
            }
        }

        //--------------------------------------------RS.GE-------------------------------------------->

        public static void operationRS(TaxInvoice oTaxInvoice, string operation, int docEntry, int seqNum, DateTime DeclDate, out string errorText, out string errorTextWb, out string errorTextGoods)
        {
            errorText = null;
            errorTextWb = null;
            errorTextGoods = null;

            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralData oGeneralDataCorr = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            oCompanyService = Program.oCompany.GetCompanyService();
            oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXS_D");
            oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            oGeneralParams.SetProperty("DocEntry", docEntry);
            oGeneralData = oGeneralService.GetByParams(oGeneralParams);

            string status = oGeneralData.GetProperty("U_status");
            bool corrInv = oGeneralData.GetProperty("U_corrInv") == "N" ? false : true;
            string declNumber = oGeneralData.GetProperty("U_declNumber");


            if ((status == "confirmed" || status == "correctionConfirmed") && (operation != "updateStatus" && operation != "cancel" && operation != "addToTheDeclaration"))
            {
                errorText = BDOSResources.getTranslate("TaxInvoiceAlreadyConfirmed");
                return;
            }

            if (corrInv)
            {
                int corrDoc = oGeneralData.GetProperty("U_corrDoc");

                //კორექტირების ა/ფ - ის მიღება
                oGeneralParams.SetProperty("DocEntry", corrDoc);
                try
                {
                    oGeneralDataCorr = oGeneralService.GetByParams(oGeneralParams);
                }
                catch
                {
                }
            }

            if (operation == "save") //შენახვა
            {
                if (corrInv == false)
                {
                    if (string.IsNullOrEmpty(status) || status == "empty")
                    {
                        save_invoice(oTaxInvoice, oGeneralData, out errorText, out errorTextWb, out errorTextGoods);
                    }
                    else
                    {
                        update_invoice(oTaxInvoice, oGeneralData, null, operation, out errorText, out errorTextWb, out errorTextGoods);
                    }
                    oGeneralService.Update(oGeneralData);
                }
                else
                {
                    if (string.IsNullOrEmpty(status) || status == "empty")
                    {
                        k_invoice(oTaxInvoice, oGeneralData, oGeneralDataCorr, out errorText, out errorTextWb, out errorTextGoods);
                    }
                    else
                    {
                        update_invoice(oTaxInvoice, oGeneralData, oGeneralDataCorr, operation, out errorText, out errorTextWb, out errorTextGoods);
                    }
                    oGeneralService.Update(oGeneralData);
                }
            }
            else if (operation == "send") //გადაგზავნა
            {
                if (string.IsNullOrEmpty(status) || status == "empty")
                {
                    if (corrInv == false)
                    {
                        save_invoice(oTaxInvoice, oGeneralData, out errorText, out errorTextWb, out errorTextGoods);
                    }
                    else
                    {
                        k_invoice(oTaxInvoice, oGeneralData, oGeneralDataCorr, out errorText, out errorTextWb, out errorTextGoods);
                    }
                }
                else
                {
                    get_invoice(oTaxInvoice, oGeneralData, operation, out errorText);
                    status = oGeneralData.GetProperty("U_status");

                    if (corrInv == false && (status == "incompleteShipped" || status == "denied" || status == "created" || status == "shipped"))
                    {
                        update_invoice(oTaxInvoice, oGeneralData, null, operation, out errorText, out errorTextWb, out errorTextGoods);
                    }
                    else if (corrInv && (status == "incompleteShipped" || status == "denied" || status == "correctionCreated" || status == "correctionShipped"))
                    {
                        update_invoice(oTaxInvoice, oGeneralData, oGeneralDataCorr, operation, out errorText, out errorTextWb, out errorTextGoods);
                    }
                }

                oGeneralService.Update(oGeneralData);
                status = oGeneralData.GetProperty("U_status");

                if (status == "created" || status == "correctionCreated" || status == "denied")
                {
                    send_invoice(oTaxInvoice, oGeneralData, operation, out errorText);
                    if (errorText == null)
                    {
                        get_invoice(oTaxInvoice, oGeneralData, operation, out errorText);
                        oGeneralService.Update(oGeneralData);
                    }
                }
                else if (status != "shipped" && status != "correctionShipped")
                {
                    if (errorText == null)
                    {
                        errorText = BDOSResources.getTranslate("TaxInvoiceShouldBeCreatedCorrectedOrDeclined");
                    }
                }
            }
            else if (operation == "updateStatus") //სტატუსების განახლება
            {
                get_invoice(oTaxInvoice, oGeneralData, operation, out errorText);
                if (errorText == null)
                {
                    oGeneralService.Update(oGeneralData);
                }
            }
            else if (operation == "remove") //წაშლა
            {
                get_invoice(oTaxInvoice, oGeneralData, operation, out errorText);
                status = oGeneralData.GetProperty("U_status");

                if (status == "removed")
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceAlreadyDeletedOnSite");
                    oGeneralService.Update(oGeneralData);
                    return;
                }

                if (status == "created" || status == "correctionCreated" || status == "incompleteShipped" || status == "shipped" || status == "correctionShipped" || status == "denied")
                {
                    remove_invoice(oTaxInvoice, oGeneralData, out errorText);
                    if (errorText == null)
                    {
                        oGeneralService.Update(oGeneralData);
                    }
                }
                else
                {
                    errorText = BDOSResources.getTranslate("StatusRequestedForDeleting");
                }
            }
            else if (operation == "cancel") //გაუქმება
            {
                get_invoice(oTaxInvoice, oGeneralData, operation, out errorText);
                status = oGeneralData.GetProperty("U_status");
                declNumber = oGeneralData.GetProperty("U_declNumber");

                if (status == "canceled")
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceAlreadyCanceledOnSite");
                    oGeneralService.Update(oGeneralData);
                    return;
                }

                if ((status == "confirmed" || status == "correctionConfirmed") && string.IsNullOrEmpty(declNumber))
                {
                    cancel_invoice(oTaxInvoice, oGeneralData, out errorText);
                    if (errorText == null)
                    {
                        oGeneralService.Update(oGeneralData);
                    }
                }
                else
                {
                    errorText = BDOSResources.getTranslate("StatusRequestedForCanceling");
                }
            }
            else if (operation == "addToTheDeclaration") //დეკლარაციაში დამატება
            {
                if (string.IsNullOrEmpty(declNumber))
                {
                    if (status == "confirmed" || status == "correctionConfirmed" || status == "primary" || status == "corrected")
                    {
                        add_inv_to_decl(oTaxInvoice, oGeneralData, seqNum, DeclDate, out errorText);
                        if (errorText == null)
                        {
                            oGeneralService.Update(oGeneralData);
                        }
                    }
                    else
                    {
                        errorText = BDOSResources.getTranslate("StatusRequestedForDeclaring");
                    }
                }
                else
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceAlreadyDeclared");
                }
            }
        }

        /// <summary>შენახვა</summary>
        private static void save_invoice(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, out string errorText, out string errorTextWb, out string errorTextGoods)
        {
            errorText = null;
            errorTextWb = null;
            errorTextGoods = null;

            try
            {
                string status = oGeneralData.GetProperty("U_status");
                string cardCodeT = oGeneralData.GetProperty("U_cardCodeT");
                bool diplomat = false;
                int buyer_un_id = oTaxInvoice.get_un_id_from_tin(cardCodeT, out diplomat, out errorText);

                if (buyer_un_id == 0)
                {
                    errorText = BDOSResources.getTranslate("CannotObtainUIDBy") + " " + cardCodeT + errorText;
                    return;
                }

                string invID_st = oGeneralData.GetProperty("U_invID");
                int inv_ID = Convert.ToInt32((string.IsNullOrEmpty(invID_st) ? "0" : invID_st));
                DateTime operation_date = oGeneralData.GetProperty("U_opDate");
                bool corrInv = oGeneralData.GetProperty("U_corrInv") == "N" ? false : true;
                string overhead_no = "";
                int b_s_user_id = 0;

                operation_date = DateTime.TryParse(operation_date.ToString("yyyyMMdd") == "18991230" ? "" : operation_date.ToString(), out operation_date) == false ? DateTime.Today : operation_date;
                bool response;

                string downPaymnt = oGeneralData.GetProperty("U_downPaymnt");
                if (downPaymnt == "Y") //საკომპენსაციო თანხის (ავანსის) ანგარიშ-ფაქტურები
                {
                    response = oTaxInvoice.save_invoice_a(ref inv_ID, operation_date, buyer_un_id, overhead_no, b_s_user_id, out errorText);
                }
                else
                {
                    response = oTaxInvoice.save_invoice(ref inv_ID, operation_date, buyer_un_id, overhead_no, b_s_user_id, out errorText);
                }

                if (response == false)
                {
                    errorText = BDOSResources.getTranslate("CreateNotSucceededSellerWoutVAT") + errorText;
                    return;
                }

                status = "0"; //შექმნილი

                //ზედნადებების ცხრილის დამატება
                response = save_ntos_invoices_inv_nos(oTaxInvoice, oGeneralData, null, inv_ID, out errorTextWb);
                if (response == false)
                {
                    status = "9"; //არასრულად შექმნილი
                }

                //საქონლის ცხრილის დამატება
                response = save_invoice_desc(oTaxInvoice, oGeneralData, inv_ID, out errorTextGoods);
                if (response == false)
                {
                    status = "9"; //არასრულად შექმნილი
                }

                bool refInv = false;
                status = getStatusValueByStatusNumber(status.ToString(), corrInv, refInv);
                oGeneralData.SetProperty("U_status", status);
                oGeneralData.SetProperty("U_invID", inv_ID.ToString());
                oGeneralData.SetProperty("U_opDate", operation_date);
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

        /// <summary>შენახვა (კორექტირების ა/ფ)</summary>
        private static void k_invoice(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, SAPbobsCOM.GeneralData oGeneralDataCorr, out string errorText, out string errorTextWb, out string errorTextGoods)
        {
            errorText = null;
            errorTextWb = null;
            errorTextGoods = null;

            try
            {
                string status = oGeneralData.GetProperty("U_status");
                string cardCodeT = oGeneralData.GetProperty("U_cardCodeT");
                bool diplomat = false;
                int buyer_un_id = oTaxInvoice.get_un_id_from_tin(cardCodeT, out diplomat, out errorText);

                if (buyer_un_id == 0)
                {
                    errorText = BDOSResources.getTranslate("CannotObtainUIDBy") + cardCodeT + errorText;
                    return;
                }

                string k_invID_st = oGeneralDataCorr.GetProperty("U_invID");
                int k_inv_ID = Convert.ToInt32((string.IsNullOrEmpty(k_invID_st) ? "0" : k_invID_st));
                DateTime operation_date = oGeneralData.GetProperty("U_opDate");
                bool corrInv = oGeneralData.GetProperty("U_corrInv") == "N" ? false : true;
                string corrType = oGeneralData.GetProperty("U_corrType");
                if (string.IsNullOrEmpty(corrType))
                {
                    errorText = BDOSResources.getTranslate("CorrectionReasonNotIndicated");
                    return;
                }
                int k_id = 0; //დაბრუნდება ა/ფ ის ID
                int k_type = Convert.ToInt32(corrType);
                string overhead_no = "";
                int b_s_user_id = 0;

                operation_date = DateTime.TryParse(operation_date.ToString("yyyyMMdd") == "18991230" ? "" : operation_date.ToString(), out operation_date) == false ? DateTime.Today : operation_date;

                bool response = oTaxInvoice.k_invoice(k_inv_ID, k_type, out k_id, out errorText);

                if (response == false)
                {
                    errorText = BDOSResources.getTranslate("FailedToCorrectTaxInvoice") + "! ID : " + k_inv_ID + errorText;
                    return;
                }

                oGeneralData.SetProperty("U_invID", k_id.ToString());
                int inv_ID = k_id;

                string downPaymnt = oGeneralData.GetProperty("U_downPaymnt");
                if (downPaymnt == "Y") //საკომპენსაციო თანხის (ავანსის) ანგარიშ-ფაქტურები
                {
                    response = oTaxInvoice.save_invoice_a(ref inv_ID, operation_date, buyer_un_id, overhead_no, b_s_user_id, out errorText);
                }
                else
                {
                    response = oTaxInvoice.save_invoice(ref inv_ID, operation_date, buyer_un_id, overhead_no, b_s_user_id, out errorText);
                }

                if (response == false)
                {
                    errorText = BDOSResources.getTranslate("CreateNotSucceededSellerWoutVAT") + errorText;
                    return;
                }

                status = "4"; //შექმნილი კორექტირებული

                //ზედნადებების ცხრილის დამატება
                response = save_ntos_invoices_inv_nos(oTaxInvoice, oGeneralData, oGeneralDataCorr, inv_ID, out errorTextWb);
                if (response == false)
                {
                    status = "9"; //არასრულად შექმნილი
                }

                //საქონლის ცხრილის დამატება
                response = save_invoice_desc(oTaxInvoice, oGeneralData, inv_ID, out errorTextGoods);
                if (response == false)
                {
                    status = "9"; //არასრულად შექმნილი
                }

                bool refInv = false;
                status = getStatusValueByStatusNumber(status.ToString(), corrInv, refInv);
                oGeneralData.SetProperty("U_status", status);
                oGeneralData.SetProperty("U_invID", inv_ID.ToString());
                oGeneralData.SetProperty("U_opDate", operation_date);
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

        /// <summary>ზედნადებების ცხრილის დამატება (ჯერ იშლება)</summary>
        private static bool save_ntos_invoices_inv_nos(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, SAPbobsCOM.GeneralData oGeneralDataCorr, int inv_ID, out string errorText)
        {
            errorText = null;
            try
            {
                string downPaymnt = oGeneralData.GetProperty("U_downPaymnt");
                if (downPaymnt == "Y")
                {
                    return true;
                }

                //ზედანდებების მიბმა --->
                int baseDoc;
                string baseDocT;
                string wbNumber;
                string actDate;
                string createDate;
                DateTime overhead_dt = new DateTime();
                Dictionary<string, string> wblDocInfo = null;
                List<string> wbNumberList = new List<string>();

                if (oGeneralDataCorr == null)
                {
                    //ზედანდებების წაშლა
                    delete_ntos_invoices_inv_nos(oTaxInvoice, inv_ID, out errorText);
                    if (errorText != null)
                    {
                        return false;
                    }
                }
                else
                {
                    DataTable taxDataTableRS = oTaxInvoice.get_ntos_invoices_inv_nos(inv_ID, out errorText);
                    DataRow taxDataRowRS;
                    int countRS = taxDataTableRS.Rows.Count;

                    for (int i = 0; i < countRS; i++)
                    {
                        taxDataRowRS = taxDataTableRS.Rows[i];
                        //int id = Convert.ToInt32(taxDataRowRS["id"]); //ზედნადების ჩანაწერის უნიკალური ID
                        //int inv_id = Convert.ToInt32(taxDataRowRS["inv_id"]); //ანგარიშ-ფაქტურის უნიკალური ნომერი
                        wbNumber = taxDataRowRS["overhead_no"].ToString(); //ზედნადების ნომერი
                        overhead_dt = Convert.ToDateTime(taxDataRowRS["overhead_dt"]); //ზედნადების თარიღი
                        //string overhead_dt_str = taxDataRowRS["overhead_dt_str"].ToString(); //ზედნადების თარიღი (სტრიქონი)  

                        wbNumberList.Add(wbNumber);
                    }
                }

                for (int i = 0; i < oGeneralData.Child("BDO_TXS1").Count; i++)
                {
                    SAPbobsCOM.GeneralData InvoiceRow = oGeneralData.Child("BDO_TXS1").Item(i);
                    baseDoc = InvoiceRow.GetProperty("U_baseDoc");
                    baseDocT = InvoiceRow.GetProperty("U_baseDocT");
                    wbNumber = InvoiceRow.GetProperty("U_wbNumber");

                    if (baseDoc == 0 || string.IsNullOrEmpty(baseDocT) || string.IsNullOrEmpty(wbNumber) || wbNumberList.Contains(wbNumber))
                    {
                        continue;
                    }

                    SAPbobsCOM.Documents oDocument;

                    if (baseDocT == "ARInvoice")
                    {
                        oDocument = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                        oDocument.GetByKey(baseDoc); //13
                        wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(baseDoc, "13", out errorText);
                    }
                    else if (baseDocT == "ARCreditNote")
                    {
                        oDocument = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                        oDocument.GetByKey(baseDoc); //14

                        if (oDocument.UserFields.Fields.Item("U_BDO_CNTp").Value == "1") //დაბრუნების ტიპია
                        {
                            wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(baseDoc, "14", out errorText);
                        }
                    }

                    else if (baseDocT == "ARCorrectionInvoice")
                    {
                        oDocument = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                        oDocument.GetByKey(baseDoc);

                        if (oDocument.UserFields.Fields.Item("U_BDOSCITp").Value == "1") //დაბრუნების ტიპია
                        {
                            wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(baseDoc, "165", out errorText);
                        }
                    }

                    if (wblDocInfo != null)
                    {
                        actDate = wblDocInfo["actDate"];
                        createDate = wblDocInfo["CreateDate"];

                        if (actDate == "30.12.1899 0:00:00")
                        {
                            DateTime.TryParse(actDate, out overhead_dt);
                        }
                        else
                        {
                            DateTime.TryParse(createDate, out overhead_dt);
                        }

                        bool response = oTaxInvoice.save_ntos_invoices_inv_nos(inv_ID, wbNumber, overhead_dt, out errorText);
                        if (response == false)
                        {
                            errorText = (errorText == null ? "" : errorText + "\n") + BDOSResources.getTranslate("CantFixWBInTaxInvoice") + " " + wbNumber + ". " + errorText;
                        }
                        else
                        {
                            wbNumberList.Add(wbNumber);
                        }
                    }
                    wblDocInfo = null;
                }

                if (errorText != null)
                {
                    return false;
                }
                else
                {
                    return true;
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
            //ზედანდებების მიბმა <---
        }

        /// <summary>საქონლის ცხრილის დამატება (ჯერ იშლება)</summary>
        private static bool save_invoice_desc(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, int inv_ID, out string errorText)
        {
            errorText = null;
            try
            {
                bool corrInv = oGeneralData.GetProperty("U_corrInv") == "N" ? false : true;
                int docEntry = Convert.ToInt32(oGeneralData.GetProperty("DocEntry"));
                //ცხრილური ნაწილის წაშლა --->
                delete_invoice_desc(oTaxInvoice, inv_ID, out errorText);
                if (errorText != null)
                {
                    return false;
                }
                //ცხრილური ნაწილის წაშლა <---             

                //ცხრილური ნაწილის დამატება --->
                int baseDoc;
                string baseDocT;

                int count = oGeneralData.Child("BDO_TXS1").Count;
                if (count > 0)
                {
                    string downPaymnt = oGeneralData.GetProperty("U_downPaymnt");
                    string query;

                    if (downPaymnt == "N")
                    {
                        List<int> docEntryARInvoiceList = new List<int>(); // A/R Invoice - ს docEntry
                        List<int> docEntryARCreditNoteList = new List<int>(); // A/R Credit Note - ს docEntry
                        List<int> docEntryARCorrectionInvoiceList = new List<int>(); // A/R Correction Invoice - ს docEntry

                        for (int i = 0; i < oGeneralData.Child("BDO_TXS1").Count; i++)
                        {
                            SAPbobsCOM.GeneralData InvoiceRow = oGeneralData.Child("BDO_TXS1").Item(i);
                            baseDoc = InvoiceRow.GetProperty("U_baseDoc");
                            baseDocT = InvoiceRow.GetProperty("U_baseDocT");

                            if (baseDoc == 0 || string.IsNullOrEmpty(baseDocT))
                            {
                                continue;
                            }
                            if (baseDocT == "ARInvoice") //A/R Invoice
                            {
                                docEntryARInvoiceList.Add(baseDoc);
                            }
                            else if (baseDocT == "ARCreditNote") //A/R Credit Note
                            {
                                docEntryARCreditNoteList.Add(baseDoc);
                            }
                            else if (baseDocT == "ARCorrectionInvoice") //A/R Correction Invoice
                            {
                                docEntryARCorrectionInvoiceList.Add(baseDoc);
                            }
                        }
                        if (corrInv)
                        {
                            errorText = getAllConnectedDoc(docEntry, ref docEntryARInvoiceList, ref docEntryARCreditNoteList, ref docEntryARCorrectionInvoiceList);
                            if (string.IsNullOrEmpty(errorText) == false)
                                return false;
                        }

                        if (docEntryARInvoiceList.Count == 0 && docEntryARCreditNoteList.Count == 0 && docEntryARCorrectionInvoiceList.Count == 0)
                        {
                            errorText = BDOSResources.getTranslate("TaxInvoiceTableEmpty");
                            return false;
                        }
                        query = getQueryArrayGoodsARInvoiceARCreditNoteARCorrectionInvoiceType(docEntryARInvoiceList, docEntryARCreditNoteList, docEntryARCorrectionInvoiceList, inv_ID);
                    }
                    else
                    {
                        List<int> docEntryARDownPaymentRequestList = new List<int>();

                        for (int i = 0; i < oGeneralData.Child("BDO_TXS1").Count; i++)
                        {
                            SAPbobsCOM.GeneralData InvoiceRow = oGeneralData.Child("BDO_TXS1").Item(i);
                            baseDoc = InvoiceRow.GetProperty("U_baseDoc");
                            baseDocT = InvoiceRow.GetProperty("U_baseDocT");

                            if (baseDoc == 0 || string.IsNullOrEmpty(baseDocT))
                            {
                                continue;
                            }
                            if (baseDocT == "ARDownPaymentVAT") //A/R Down Payment VAT
                            {
                                docEntryARDownPaymentRequestList.Add(baseDoc);
                            }
                        }
                        if (docEntryARDownPaymentRequestList.Count == 0)
                        {
                            errorText = BDOSResources.getTranslate("TaxInvoiceTableEmpty");
                            return false;
                        }
                        query = getQueryArrayGoodsARDownPaymentVATType(string.Join("','", docEntryARDownPaymentRequestList), inv_ID);
                    }

                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery(query);
                    int countRS = oRecordSet.RecordCount;

                    if (countRS == 0)
                    {
                        errorText = BDOSResources.getTranslate("TaxInvoiceTableEmpty");
                        return false;
                    }

                    while (!oRecordSet.EoF)
                    {
                        int id = 0; //ანგარიშ-ფაქტურის საქონლის მონაცემის უნიკალური ნომერი
                        int inv_id = inv_ID; //ანგარიშ-ფაქტურის უნიკალური ნომერი
                        string goods = oRecordSet.Fields.Item("W_NAME").Value.ToString(); //საქონლის დასახელება
                        string g_unit = oRecordSet.Fields.Item("InvntItem").Value == "N" ? "მომსახურება" : oRecordSet.Fields.Item("UNIT_TXT").Value.ToString(); //საქონლის ერთეული
                        if (g_unit == "")
                        {
                            g_unit = "სხვა";
                        }
                        decimal g_number = Convert.ToDecimal(oRecordSet.Fields.Item("QUANTITY").Value); //რაოდენობა
                        decimal full_amount = Convert.ToDecimal(oRecordSet.Fields.Item("AMOUNT").Value); //თანხა დღგ-ის და აქციზის ჩათვლლით
                        decimal drg_amount = Convert.ToDecimal(oRecordSet.Fields.Item("LineVat").Value); //დღგ
                        decimal aqcizi_amount = 0; //აქციზი
                        int akcis_id = 0; //აქციზური საქონლის კოდი  

                        bool response = oTaxInvoice.save_invoice_desc(id, inv_ID, goods, g_unit, g_number, full_amount, drg_amount, aqcizi_amount, akcis_id, out errorText);

                        if (response == false)
                        {
                            errorText = (errorText == null ? "" : errorText + "\n") + BDOSResources.getTranslate("CantFixGoodsItemInTaxInvoice") + goods + "\". " + errorText;
                        }

                        oRecordSet.MoveNext();
                    }
                    if (errorText != null)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceTableEmpty");
                    return false;
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
            //ცხრილური ნაწილის დამატება <---
        }

        /// <summary>რეკვიზიტების განახლება</summary>
        private static void get_invoice(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, string operation, out string errorText)
        {
            errorText = null;

            try
            {
                string invID = oGeneralData.GetProperty("U_invID");
                if (string.IsNullOrEmpty(invID))
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
                    string f_number = responseDictionary["f_number"] == null ? "" : responseDictionary["f_number"].ToString();
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

                    string statusDoc = oGeneralData.GetProperty("U_status");
                    bool refInv = false;
                    string status = getStatusValueByStatusNumber(statusRS.ToString(), corrInv, refInv);
                    string invoice_no = oGeneralData.GetProperty("U_number");

                    if (string.IsNullOrEmpty(invoice_no))
                    {
                        invoice_no = f_number;
                    }
                    //else
                    //{
                    //    invoice_no = f_number == null ? "" : invoice_no;
                    //}

                    if ((statusDoc == "shipped" && status == "created") || (statusDoc == "correctionShipped" && status == "correctionCreated"))
                    {
                        status = "denied";
                    }

                    oGeneralData.SetProperty("U_status", status);
                    oGeneralData.SetProperty("U_opDate", operation_dt);
                    oGeneralData.SetProperty("U_sentDate", reg_dt);
                    oGeneralData.SetProperty("U_number", invoice_no);
                    oGeneralData.SetProperty("U_series", f_series);
                    oGeneralData.SetProperty("U_declNumber", seq_num_s);

                    //მგონი დასაკომენტარებელია???!!!
                    //DateTime opDate = oGeneralData.GetProperty("U_opDate");

                    //if (operation != "send" && operation != "save" && operation != "remove" && operation != "cancel") // როცა ვქმნით, ვაგზავნით, ვშლით ან ვაუქმებთ აქ არ უნდა შემოვიდეს ა/ფ-ის ნომერი ახალი მინიჭებული აქვს ან არ აქვს და არ არის საჭიროება აქ შევიდეს.
                    //{
                    //    if (string.IsNullOrEmpty(invoice_no) == false && opDate != new DateTime())
                    //    {
                    //        get_seller_invoices( oTaxInvoice, oGeneralData, null, out errorText);
                    //    }
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

        public static string getQueryArrayGoodsARInvoiceARCreditNoteARCorrectionInvoiceType(List<int> docEntryARInvoiceList, List<int> docEntryARCreditNoteList, List<int> docEntryARCorrectionInvoiceList, int inv_ID)
        {
            if (docEntryARInvoiceList.Count == 0)
                docEntryARInvoiceList.Add(-99);
            if (docEntryARCreditNoteList.Count == 0)
                docEntryARCreditNoteList.Add(-99);
            if (docEntryARCorrectionInvoiceList.Count == 0)
                docEntryARCorrectionInvoiceList.Add(-99);

            //""BDO_RSUOM"".""U_RSCode"" AS ""UNIT_ID"",
            //     ""OITM"".""InvntryUom"" AS ""UNIT_TXT"",

            //CASE WHEN ""BDO_RSUOM"".""U_RSCode"" is null THEN '99' ELSE ""BDO_RSUOM"".""U_RSCode"" END AS ""UNIT_ID"", 
            //CASE WHEN ""OITM"".""InvntryUom"" = '' THEN 'სხვა' ELSE ""OITM"".""InvntryUom"" END  AS ""UNIT_TXT"",


            string query = @"SELECT
	                         ""W_NAME"",
	                         ""InvntItem"",
	                         ""UNIT_TXT"",
	                         SUM(""QUANTITY"") AS ""QUANTITY"",
	                         SUM(""AMOUNT"") AS ""AMOUNT"",
	                         SUM(""LineVat"") AS ""LineVat"" 
                        FROM (SELECT
	                         '" + inv_ID + @"' AS ""ID"",
	                         ""MNTB"".""LineNum"" AS ""LineNum"",
	                         ""MNTB"".""DocEntry"" AS ""DocEntry"",
	                         ""MNTB"".""ItemCode"" AS ""ItemCode"",
	                         ""OITM"".""CodeBars"" AS ""CodeBars"",
	                         ""OITM"".""SWW"" AS ""AdditionalIdentifier"",
	                         ""MNTB"".""Dscription"" AS ""W_NAME"",
                             CASE WHEN ""BDO_RSUOM"".""U_RSCode"" is null THEN '99' ELSE ""BDO_RSUOM"".""U_RSCode"" END AS ""UNIT_ID"", 
                             CASE WHEN ""MNTB"".""unitMsr"" = '' or ""MNTB"".""U_BDOSSrvDsc"" <> '' THEN 'სხვა' ELSE ""MNTB"".""unitMsr"" END  AS ""UNIT_TXT"",
	                         ""MNTB"".""VatPrcnt"" AS ""VAT_TYPE"",
	                         ""MNTB"".""VatGroup""AS ""VatGroup"",
	                         '0' AS ""A_ID"",
	                         SUM(""MNTB"".""Quantity"") AS ""QUANTITY"",
	                         SUM(""MNTB"".""GTotal"") AS ""AMOUNT"",
	                         CASE WHEN SUM(""MNTB"".""Quantity"") = 0 
	                        THEN 0 
	                        ELSE SUM(""MNTB"".""GTotal"")/SUM(""MNTB"".""Quantity"") 
	                        END AS ""PRICE"",
	                         SUM(""MNTB"".""LineVat"") AS ""LineVat"",
	                         ""MNTB"".""ItemType"" AS ""ItemType"",
	                         ""MNTB"".""InvntItem"" AS ""InvntItem"" 
	                        FROM (SELECT
	                         ""ORIN"".""U_BDOSSrvDsc"" AS ""U_BDOSSrvDsc"",
	                         ""RIN1"".""DocEntry"" AS ""DocEntry"",
	                         ""RIN1"".""BaseEntry"" AS ""BaseEntry"",
	                         ""RIN1"".""BaseLine"" AS ""LineNum"",
--""RIN1"".""ItemCode"",
                            (CASE WHEN ""ORIN"".""U_BDOSSrvDsc"" is null 
	                        		THEN ""RIN1"".""ItemCode"" 
	                        		ELSE '' 
	                        		END)  AS ""ItemCode"",
--""RIN1"".""Dscription"",
                            (CASE WHEN ""ORIN"".""U_BDOSSrvDsc"" is null 
	                        		THEN ""RIN1"".""Dscription"" 
	                        		ELSE ""ORIN"".""U_BDOSSrvDsc"" 
	                        		END)  AS ""Dscription"",
                             ""RIN1"".""unitMsr"",
                             ""RIN1"".""Quantity"" * (-1) * (CASE WHEN ""RIN1"".""NoInvtryMv"" = 'Y' 
	                        		THEN 0 
	                        		ELSE 1 
	                        		END) * ""RIN1"".""NumPerMsr"" AS ""Quantity"",
	                         ""RIN1"".""GTotal"" * (-1) AS ""GTotal"",
	                         ""RIN1"".""VatPrcnt"",
	                         ""RIN1"".""VatGroup"",
	                         ""RIN1"".""LineVat"" * (-1) AS ""LineVat"",
--""OITM"".""ItemType"",
                             CASE WHEN ""ORIN"".""U_BDOSSrvDsc"" is null THEN ""OITM"".""ItemType"" ELSE 'L' END AS ""ItemType"", 
--""OITM"".""InvntItem"" 
                             CASE WHEN ""ORIN"".""U_BDOSSrvDsc"" is null THEN ""OITM"".""InvntItem"" ELSE 'N' END AS ""InvntItem""
		                        FROM ""RIN1"" 
		                        INNER JOIN ""ORIN"" ON ""RIN1"".""DocEntry"" = ""ORIN"".""DocEntry"" 
		                        LEFT JOIN ""OITM"" AS ""OITM"" ON ""RIN1"".""ItemCode"" = ""OITM"".""ItemCode"" 
		                        WHERE ""RIN1"".""DocEntry"" IN (" + string.Join(",", docEntryARCreditNoteList) + @") 
		                        AND ""RIN1"".""BaseEntry"" IN (" + string.Join(",", docEntryARInvoiceList) + @") 
		                        AND ""RIN1"".""TargetType"" < 0 

                                UNION ALL
                                
                                 SELECT
	                         ""OCSI"".""U_BDOSSrvDsc"" AS ""U_BDOSSrvDsc"",
	                         ""CSI1"".""DocEntry"" AS ""DocEntry"",
	                         ""CSI1"".""BaseEntry"" AS ""BaseEntry"",
	                         ""CSI1"".""BaseLine"" AS ""LineNum"",
--""CSI1"".""ItemCode"",
                            (CASE WHEN ""OCSI"".""U_BDOSSrvDsc"" is null 
	                        		THEN ""CSI1"".""ItemCode"" 
	                        		ELSE '' 
	                        		END)  AS ""ItemCode"",
--""CSI1"".""Dscription"",
                            (CASE WHEN ""OCSI"".""U_BDOSSrvDsc"" is null 
	                        		THEN ""CSI1"".""Dscription"" 
	                        		ELSE ""OCSI"".""U_BDOSSrvDsc"" 
	                        		END)  AS ""Dscription"",
                             ""CSI1"".""unitMsr"",
                             ""CSI1"".""Quantity"" * (CASE WHEN ""CSI1"".""NoInvtryMv"" = 'Y' 
	                        		THEN 0 
	                        		ELSE 1 
	                        		END) * ""CSI1"".""NumPerMsr"" AS ""Quantity"",
	                         ""CSI1"".""GTotal"" AS ""GTotal"",
	                         ""CSI1"".""VatPrcnt"",
	                         ""CSI1"".""VatGroup"",
	                         ""CSI1"".""LineVat"" AS ""LineVat"",
--""OITM"".""ItemType"",
                             CASE WHEN ""OCSI"".""U_BDOSSrvDsc"" is null THEN ""OITM"".""ItemType"" ELSE 'L' END AS ""ItemType"", 
--""OITM"".""InvntItem"" 
                             CASE WHEN ""OCSI"".""U_BDOSSrvDsc"" is null THEN ""OITM"".""InvntItem"" ELSE 'N' END AS ""InvntItem""
		                        FROM ""CSI1"" 
		                        INNER JOIN ""OCSI"" ON ""CSI1"".""DocEntry"" = ""OCSI"".""DocEntry"" 
		                        LEFT JOIN ""OITM"" AS ""OITM"" ON ""CSI1"".""ItemCode"" = ""OITM"".""ItemCode"" 
		                        WHERE ""CSI1"".""DocEntry"" IN (" + string.Join(",", docEntryARCorrectionInvoiceList) + @") 
		                        AND ""CSI1"".""BaseEntry"" IN (" + string.Join(",", docEntryARInvoiceList) + @") 
		                        AND ""CSI1"".""TargetType"" < 0 ) AS ""MNTB"" 
	                        LEFT JOIN ""OITM"" AS ""OITM"" ON ""MNTB"".""ItemCode"" = ""OITM"".""ItemCode"" 
	                        LEFT JOIN ""OUOM"" AS ""OUOM"" ON ""MNTB"".""unitMsr"" = ""OUOM"".""UomName"" 
	                        LEFT JOIN ""@BDO_RSUOM"" AS ""BDO_RSUOM"" ON ""OUOM"".""UomEntry"" = ""BDO_RSUOM"".""U_UomEntry"" 
	                        GROUP BY ""MNTB"".""U_BDOSSrvDsc"",
	                         ""MNTB"".""DocEntry"",
	                         ""MNTB"".""LineNum"",
	                         ""MNTB"".""ItemCode"",
	                         ""MNTB"".""Dscription"",
	                         ""OITM"".""CodeBars"",
	                         ""OITM"".""SWW"",
	                         ""BDO_RSUOM"".""U_RSCode"",
	                         ""MNTB"".""unitMsr"",
	                         ""MNTB"".""VatPrcnt"",
	                         ""MNTB"".""VatGroup"",
	                         ""MNTB"".""ItemType"",
	                         ""MNTB"".""InvntItem"" 
	                        UNION ALL SELECT
	                         '" + inv_ID + @"' AS ""ID"",
	                         ""MNTB"".""LineNum"" AS ""LineNum"",
	                         ""MNTB"".""DocEntry"" AS ""DocEntry"",
	                         ""MNTB"".""ItemCode"" AS ""ItemCode"",
	                         ""OITM"".""CodeBars"" AS ""CodeBars"",
	                         ""OITM"".""SWW"" AS ""AdditionalIdentifier"",
	                         ""MNTB"".""Dscription"" AS ""W_NAME"",
                             CASE WHEN ""BDO_RSUOM"".""U_RSCode"" is null THEN '99' ELSE ""BDO_RSUOM"".""U_RSCode"" END AS ""UNIT_ID"", 
                             CASE WHEN ""MNTB"".""unitMsr"" = '' or ""MNTB"".""U_BDOSSrvDsc"" <> '' THEN 'სხვა' ELSE ""MNTB"".""unitMsr"" END  AS ""UNIT_TXT"",
	                         ""MNTB"".""VatPrcnt"" AS ""VAT_TYPE"",
	                         ""MNTB"".""VatGroup""AS ""VatGroup"",
	                         '0' AS ""A_ID"",
	                         SUM(""MNTB"".""Quantity"") AS ""QUANTITY"",
	                         SUM(""MNTB"".""GTotal"") AS ""AMOUNT"",
	                         CASE WHEN SUM(""MNTB"".""Quantity"") = 0 
	                        THEN 0 
	                        ELSE SUM(""MNTB"".""GTotal"")/SUM(""MNTB"".""Quantity"") 
	                        END AS ""PRICE"",
	                         SUM(""MNTB"".""LineVat"") AS ""LineVat"",
                        	 ""MNTB"".""ItemType"" AS ""ItemType"",
                        	 ""MNTB"".""InvntItem"" AS ""InvntItem"" 
                        	FROM (SELECT
                        	 ""OINV"".""U_BDOSSrvDsc"",
                        	 ""INV1"".""DocEntry"",
                        	 ""INV1"".""LineNum"",
--""INV1"".""ItemCode"",
                            (CASE WHEN ""OINV"".""U_BDOSSrvDsc"" is null 
	                        		THEN ""INV1"".""ItemCode"" 
	                        		ELSE '' 
	                        		END)  AS ""ItemCode"",
--""INV1"".""Dscription"",
                            (CASE WHEN ""OINV"".""U_BDOSSrvDsc"" is null 
	                        		THEN ""INV1"".""Dscription"" 
	                        		ELSE ""OINV"".""U_BDOSSrvDsc"" 
	                        		END)  AS ""Dscription"",
                             ""INV1"".""unitMsr"",
                        	 (CASE WHEN ""INV1"".""ItemCode"" is null 
                        			THEN 1 
                        			ELSE ""INV1"".""Quantity"" 
                        			END) * (CASE WHEN ""INV1"".""ItemCode"" is null 
                        			THEN 1 
	                        		ELSE ""INV1"".""NumPerMsr"" 
	                        		END) AS ""Quantity"",
	                         ""INV1"".""GTotal"",
	                         ""INV1"".""VatPrcnt"",
	                         ""INV1"".""VatGroup"",
	                         ""INV1"".""LineVat"",
--""OITM"".""ItemType"",
                             CASE WHEN ""OINV"".""U_BDOSSrvDsc"" is null THEN ""OITM"".""ItemType"" ELSE 'L' END AS ""ItemType"", 
--""OITM"".""InvntItem"" 
                             CASE WHEN ""OINV"".""U_BDOSSrvDsc"" is null THEN ""OITM"".""InvntItem"" ELSE 'N' END AS ""InvntItem""
	                        	FROM ""INV1"" 
	                        	LEFT JOIN ""OINV"" AS ""OINV"" ON ""INV1"".""DocEntry"" = ""OINV"".""DocEntry"" 
	                        	LEFT JOIN ""OITM"" AS ""OITM"" ON ""INV1"".""ItemCode"" = ""OITM"".""ItemCode"" 
	                        	WHERE ""INV1"".""DocEntry"" IN (" + string.Join(",", docEntryARInvoiceList) + @") ) AS ""MNTB"" 
	                        LEFT JOIN ""OITM"" AS ""OITM"" ON ""MNTB"".""ItemCode"" = ""OITM"".""ItemCode"" 
	                        LEFT JOIN ""OUOM"" AS ""OUOM"" ON ""MNTB"".""unitMsr"" = ""OUOM"".""UomName"" 
	                        LEFT JOIN ""@BDO_RSUOM"" AS ""BDO_RSUOM"" ON ""OUOM"".""UomEntry"" = ""BDO_RSUOM"".""U_UomEntry"" 
	                        GROUP BY ""MNTB"".""U_BDOSSrvDsc"",
	                         ""MNTB"".""DocEntry"",
	                         ""MNTB"".""LineNum"",
	                         ""MNTB"".""ItemCode"",
	                         ""MNTB"".""Dscription"",
	                         ""OITM"".""CodeBars"",
	                         ""OITM"".""SWW"",
	                         ""BDO_RSUOM"".""U_RSCode"",
	                         ""MNTB"".""unitMsr"",
	                         ""MNTB"".""VatPrcnt"",
	                         ""MNTB"".""VatGroup"",
	                         ""MNTB"".""ItemType"",
	                         ""MNTB"".""InvntItem"" HAVING SUM(""MNTB"".""Quantity"") > 0 ) AS ""MNTB_DATA""
                        GROUP BY ""W_NAME"",
                        	 ""InvntItem"",
	                         ""UNIT_TXT""";

            return query;
        }

        private static string getQueryArrayGoodsARDownPaymentVATType(string baseDocEntry, int inv_ID)
        {
            string query = @"SELECT
            	 ""MNTB"".""ID"" AS ""ID"",
            	 ""MNTB"".""DocEntry"" AS ""DocEntry"",
            	 ""MNTB"".""LineId"" AS ""LineNum"",
            	 ""MNTB"".""U_ItemCode"" AS ""ItemCode"",
            	 ""MNTB"".""U_Dscptn"" AS ""W_NAME"",
            	 SUM(""MNTB"".""QUANTITY"") AS ""QUANTITY"",
            	 SUM(""MNTB"".""GTotal"") AS ""AMOUNT"",
            	 SUM(""MNTB"".""LineVat"") AS ""LineVat"",
            	 CASE WHEN SUM(""MNTB"".""QUANTITY"") = 0 
            THEN 0 
            ELSE SUM(""MNTB"".""GTotal"")/SUM(""MNTB"".""QUANTITY"") 
            END AS ""PRICE"",
            	 ""MNTB"".""VatPrcnt"" AS ""VAT_TYPE"",
            	 ""MNTB"".""U_VatGrp"" AS ""VatGroup"",
            	 ""MNTB"".""ItemType"" AS ""ItemType"",
            	 ""MNTB"".""InvntItem"" AS ""InvntItem"",
            	 ""MNTB"".""CodeBars"" AS ""CodeBars"",
            	 ""MNTB"".""SWW"" AS ""AdditionalIdentifier"",
                 CASE WHEN ""MNTB"".""U_RSCode"" is null THEN '99' ELSE ""MNTB"".""U_RSCode"" END AS ""UNIT_ID"", 
                 CASE WHEN ""MNTB"".""InvntryUom"" = '' THEN 'სხვა' ELSE ""MNTB"".""InvntryUom"" END  AS ""UNIT_TXT"",            	 
            	 ""MNTB"".""A_ID"" AS ""A_ID"" 
            FROM (SELECT " + " " +
                 @"'" + inv_ID + @"' AS ""ID"",
            	 ""@BDOSRDV1"".""DocEntry"",
            	 ""@BDOSRDV1"".""LineId"",
            	 ""@BDOSRDV1"".""U_ItemCode"",
            	 ""@BDOSRDV1"".""U_Dscptn"",
            	 ""@BDOSRDV1"".""U_Qnty"" AS ""QUANTITY"",
            	 (""@BDOSRDV1"".""U_GrsAmnt"") AS ""GTotal"",
            	 18 AS ""VatPrcnt"",
            	 ""@BDOSRDV1"".""U_VatGrp"",
            	 (""@BDOSRDV1"".""U_VatAmount"") AS  ""LineVat"",
            	 ""OITM"".""ItemType"",
            	 ""OITM"".""InvntItem"",
            	 ""OITM"".""InvntryUom"",
            	 ""OITM"".""CodeBars"",
            	 ""OITM"".""SWW"",
            	 ""BDO_RSUOM"".""U_RSCode"",
            	 '0' AS ""A_ID"" 
            	FROM ""@BDOSRDV1"" 
            	LEFT JOIN ""OITM"" AS ""OITM"" ON ""@BDOSRDV1"".""U_ItemCode"" = ""OITM"".""ItemCode"" 
            	LEFT JOIN ""OUOM"" AS ""OUOM"" ON ""OITM"".""InvntryUom"" = ""OUOM"".""UomName"" 
            	LEFT JOIN ""@BDO_RSUOM"" AS ""BDO_RSUOM"" ON ""OUOM"".""UomEntry"" = ""BDO_RSUOM"".""U_UomEntry"" 
            	WHERE ""@BDOSRDV1"".""DocEntry"" IN ('" + baseDocEntry + @"')) AS ""MNTB"" 
            GROUP BY ""MNTB"".""ID"",
            	 ""MNTB"".""DocEntry"",
            	 ""MNTB"".""LineId"",
            	 ""MNTB"".""U_ItemCode"",
            	 ""MNTB"".""U_Dscptn"",
            	 ""MNTB"".""VatPrcnt"",
            	 ""MNTB"".""U_VatGrp"",
            	 ""MNTB"".""ItemType"",
            	 ""MNTB"".""InvntItem"",
            	 ""MNTB"".""InvntryUom"",
            	 ""MNTB"".""CodeBars"",
            	 ""MNTB"".""SWW"",
            	 ""MNTB"".""U_RSCode"",
            	 ""MNTB"".""A_ID"" ";

            return query;
        }

        private static string getQueryArrayGoodsARDownPaymentRequestType(string baseDocEntry, int inv_ID)
        {
            string query = @"SELECT
            	 ""MNTB"".""ID"" AS ""ID"",
            	 ""MNTB"".""DocEntry"" AS ""DocEntry"",
            	 ""MNTB"".""LineNum"" AS ""LineNum"",
            	 ""MNTB"".""ItemCode"" AS ""ItemCode"",
            	 ""MNTB"".""Dscription"" AS ""W_NAME"",
            	 SUM(""MNTB"".""QUANTITY"") AS ""QUANTITY"",
            	 SUM(""MNTB"".""GTotal"") AS ""AMOUNT"",
            	 SUM(""MNTB"".""LineVat"") AS ""LineVat"",
            	 CASE WHEN SUM(""MNTB"".""QUANTITY"") = 0 
            THEN 0 
            ELSE SUM(""MNTB"".""GTotal"")/SUM(""MNTB"".""QUANTITY"") 
            END AS ""PRICE"",
            	 ""MNTB"".""VatPrcnt"" AS ""VAT_TYPE"",
            	 ""MNTB"".""VatGroup"" AS ""VatGroup"",
            	 ""MNTB"".""ItemType"" AS ""ItemType"",
            	 ""MNTB"".""InvntItem"" AS ""InvntItem"",
            	 ""MNTB"".""CodeBars"" AS ""CodeBars"",
            	 ""MNTB"".""SWW"" AS ""AdditionalIdentifier"",
                 CASE WHEN ""MNTB"".""U_RSCode"" is null THEN '99' ELSE ""MNTB"".""U_RSCode"" END AS ""UNIT_ID"", 
                 CASE WHEN ""MNTB"".""InvntryUom"" = '' THEN 'სხვა' ELSE ""MNTB"".""InvntryUom"" END  AS ""UNIT_TXT"",            	 
            	 ""MNTB"".""A_ID"" AS ""A_ID"" 
            FROM (SELECT " + " " +
                 @"'" + inv_ID + @"' AS ""ID"",
            	 ""DPI1"".""DocEntry"",
            	 ""DPI1"".""LineNum"",
            	 ""DPI1"".""ItemCode"",
            	 ""DPI1"".""Dscription"",
            	 (CASE WHEN ""DPI1"".""ItemCode"" is null 
            		THEN 1 
            		ELSE ""DPI1"".""Quantity"" 
            		END) * (CASE WHEN ""DPI1"".""ItemCode"" is null 
            		THEN 1 
            		ELSE ""DPI1"".""NumPerMsr"" 
            		END) AS ""QUANTITY"",
            	 (""DPI1"".""U_BDOSDPMAmt"") AS ""GTotal"",
            	 ""DPI1"".""VatPrcnt"",
            	 ""DPI1"".""VatGroup"",
            	 (""DPI1"".""U_BDOSDPMVat"") AS  ""LineVat"",
            	 ""OITM"".""ItemType"",
            	 ""OITM"".""InvntItem"",
            	 ""OITM"".""InvntryUom"",
            	 ""OITM"".""CodeBars"",
            	 ""OITM"".""SWW"",
            	 ""BDO_RSUOM"".""U_RSCode"",
            	 '0' AS ""A_ID"" 
            	FROM ""DPI1"" 
            	LEFT JOIN ""OITM"" AS ""OITM"" ON ""DPI1"".""ItemCode"" = ""OITM"".""ItemCode"" 
            	LEFT JOIN ""OUOM"" AS ""OUOM"" ON ""OITM"".""InvntryUom"" = ""OUOM"".""UomName"" 
            	LEFT JOIN ""@BDO_RSUOM"" AS ""BDO_RSUOM"" ON ""OUOM"".""UomEntry"" = ""BDO_RSUOM"".""U_UomEntry"" 
            	WHERE ""DPI1"".""DocEntry"" IN ('" + baseDocEntry + @"')) AS ""MNTB"" 
            GROUP BY ""MNTB"".""ID"",
            	 ""MNTB"".""DocEntry"",
            	 ""MNTB"".""LineNum"",
            	 ""MNTB"".""ItemCode"",
            	 ""MNTB"".""Dscription"",
            	 ""MNTB"".""VatPrcnt"",
            	 ""MNTB"".""VatGroup"",
            	 ""MNTB"".""ItemType"",
            	 ""MNTB"".""InvntItem"",
            	 ""MNTB"".""InvntryUom"",
            	 ""MNTB"".""CodeBars"",
            	 ""MNTB"".""SWW"",
            	 ""MNTB"".""U_RSCode"",
            	 ""MNTB"".""A_ID"" HAVING SUM(""MNTB"".""QUANTITY"") > 0";

            return query;
        }

        /// <summary>საქონლის ცხრ. ნაწილის წაშლა ფაქტურაში</summary>
        private static void delete_invoice_desc(TaxInvoice oTaxInvoice, int inv_ID, out string errorText)
        {
            DataTable taxDataTableRS = oTaxInvoice.get_invoice_desc(inv_ID, out errorText);
            DataRow taxDataRowRS;
            int countRS = taxDataTableRS.Rows.Count;

            for (int i = 0; i < countRS; i++)
            {
                taxDataRowRS = taxDataTableRS.Rows[i];
                int id = Convert.ToInt32(taxDataRowRS["id"]); //ანგარიშ-ფაქტურის საქონლის მონაცემის უნიკალური ნომერი
                int inv_id = Convert.ToInt32(taxDataRowRS["inv_id"]); //ანგარიშ-ფაქტურის უნიკალური ნომერი
                string goods = taxDataRowRS["goods"].ToString(); //საქონლის დასახელება
                //taxDataRowRS["g_unit"].ToString(); //საქონლის ერთეული
                //taxDataRowRS["g_number"].ToString(); //რაოდენობა
                //taxDataRowRS["full_amount"].ToString(); //თანხა დღგ-ის და აქციზის ჩათვლლით
                //taxDataRowRS["drg_amount"].ToString(); //დღგ
                //taxDataRowRS["aqcizi_amount"].ToString(); //აქციზი
                //taxDataRowRS["akcis_id"].ToString(); //აქციზური საქონლის კოდი
                //taxDataRowRS["sdrg_amount"].ToString(); //დღგ სტრიქონული ტიპის

                bool response = oTaxInvoice.delete_invoice_desc(id, inv_id, out errorText);
                if (response == false)
                {
                    errorText = BDOSResources.getTranslate("CantDeleteGoodsInTaxInvoice") + goods + "\". " + errorText;
                    return;
                }

                //g_number = g_number + Convert.ToDouble(TaxDeclRow.ItemArray[4], CultureInfo.InvariantCulture); //რაოდენობა
                //full_amount = full_amount + Convert.ToDouble(TaxDeclRow.ItemArray[5], CultureInfo.InvariantCulture); //თანხა დღგ-ის და აქციზის ჩათვლით
                //drg_amount = drg_amount + Convert.ToDouble(TaxDeclRow.ItemArray[6], CultureInfo.InvariantCulture); //დღგ
                //aqcizi_amount = aqcizi_amount + Convert.ToDouble(TaxDeclRow.ItemArray[7], CultureInfo.InvariantCulture); //აქციზის თანხა
            }
        }

        /// <summary>ზედნადების ცხრ. ნაწილის წაშლა ფაქტურაში</summary>
        private static void delete_ntos_invoices_inv_nos(TaxInvoice oTaxInvoice, int inv_ID, out string errorText)
        {
            DataTable taxDataTableRS = oTaxInvoice.get_ntos_invoices_inv_nos(inv_ID, out errorText);
            DataRow taxDataRowRS;
            int countRS = taxDataTableRS.Rows.Count;

            for (int i = 0; i < countRS; i++)
            {
                taxDataRowRS = taxDataTableRS.Rows[i];
                int id = Convert.ToInt32(taxDataRowRS["id"]); //ზედნადების ჩანაწერის უნიკალური ID
                int inv_id = Convert.ToInt32(taxDataRowRS["inv_id"]); //ანგარიშ-ფაქტურის უნიკალური ნომერი
                string overhead_no = taxDataRowRS["overhead_no"].ToString(); //ზედნადების ნომერი
                DateTime overhead_dt = Convert.ToDateTime(taxDataRowRS["overhead_dt"]); //ზედნადების თარიღი
                string overhead_dt_str = taxDataRowRS["overhead_dt_str"].ToString(); //ზედნადების თარიღი (სტრიქონი)

                bool response = oTaxInvoice.delete_ntos_invoices_inv_nos(id, inv_id, out errorText);
                if (response == false)
                {
                    errorText = BDOSResources.getTranslate("CantDeleteWBInTaxInvoice") + overhead_no + "\". " + errorText;
                    return;
                }
            }
        }

        /// <summary>გადაგზავნა</summary>
        private static void send_invoice(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, string operation, out string errorText)
        {
            errorText = null;

            try
            {
                string cardCodeT = oGeneralData.GetProperty("U_cardCodeT");
                string invID = oGeneralData.GetProperty("U_invID");
                bool corrInv = oGeneralData.GetProperty("U_corrInv") == "N" ? false : true;

                if (string.IsNullOrEmpty(invID))
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceIDNotFilled");
                    return;
                }

                bool diplomat = false;
                int buyer_un_id = oTaxInvoice.get_un_id_from_tin(cardCodeT, out diplomat, out errorText);

                if (buyer_un_id == 0)
                {
                    errorText = BDOSResources.getTranslate("CannotObtainUIDBy") + cardCodeT + errorText;
                    return;
                }

                int inv_ID = Convert.ToInt32(invID);

                int statusRS = corrInv == false ? 1 : 5;

                statusRS = diplomat ? 6 : statusRS;

                bool response = oTaxInvoice.change_invoice_status(inv_ID, statusRS, out errorText);
                if (response == false)
                {
                    errorText = BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("Send") + " " + BDOSResources.getTranslate("DoneWithErrors") + " : " + errorText;
                    return;
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

        /// <summary>წაშლა</summary>
        private static void remove_invoice(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, out string errorText)
        {
            errorText = null;

            try
            {
                string invID = oGeneralData.GetProperty("U_invID");
                bool corrInv = oGeneralData.GetProperty("U_corrInv") == "N" ? false : true;

                if (string.IsNullOrEmpty(invID))
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceIDNotFilled");
                    return;
                }

                int inv_ID = Convert.ToInt32(invID);
                bool response = oTaxInvoice.change_invoice_status(inv_ID, -1, out errorText);
                if (response == false)
                {
                    errorText = BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("Delete") + " " + BDOSResources.getTranslate("DoneWithErrors") + " : " + errorText;
                    return;
                }
                else
                {
                    bool refInv = false;
                    string status = getStatusValueByStatusNumber("-1", corrInv, refInv); //წაშლილი

                    oGeneralData.SetProperty("U_status", status);
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

        /// <summary>გაუქმება</summary>
        private static void cancel_invoice(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, out string errorText)
        {
            errorText = null;

            try
            {
                string invID = oGeneralData.GetProperty("U_invID");
                string status = oGeneralData.GetProperty("U_status");
                bool corrInv = oGeneralData.GetProperty("U_corrInv") == "N" ? false : true;

                if (string.IsNullOrEmpty(invID))
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceIDNotFilled");
                    return;
                }

                if (status != "confirmed" & status != "correctionConfirmed")
                {
                    errorText = BDOSResources.getTranslate("StatusRequestedForCanceling");
                    return;
                }

                int inv_ID = Convert.ToInt32(invID);
                bool response = oTaxInvoice.change_invoice_status(inv_ID, 6, out errorText);
                if (response == false)
                {
                    errorText = BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("Cancel") + " " + BDOSResources.getTranslate("DoneWithErrors") + " : " + errorText;
                    return;
                }
                else
                {
                    bool refInv = false;
                    status = getStatusValueByStatusNumber("6", corrInv, refInv); //გაუქმების პროცესში

                    oGeneralData.SetProperty("U_status", status);
                    //oGeneralData.SetProperty("U_canDate", DateTime.Today); გაუქმების თარიღი ???
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
                string invID = oGeneralData.GetProperty("U_invID");
                string status = oGeneralData.GetProperty("U_status");

                if (string.IsNullOrEmpty(invID))
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceIDNotFilled");
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

                if (response)
                {
                    oGeneralData.SetProperty("U_declNumber", seqNum.ToString());
                    oGeneralData.SetProperty("U_declDate", declDate);
                }
                else
                {
                    errorText = BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("RSAddDeclaration") + " " + BDOSResources.getTranslate("DoneWithErrors") + " : " + errorText;
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

        /// <summary>რეკვიზიტების განახლება (+ აბრუნებს დადასტურების თარიღს, აბრუნებს უარყოფილია თუ არა (წაშლილი ა/ფ არ ჩანს))</summary>
        private static void get_seller_invoices(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, DataRow taxDataRow, out string errorText)
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

                    DataTable taxDataTable = oTaxInvoice.get_seller_invoices(s_dt, e_dt, op_s_dt, op_e_dt, "000" + invoice_no, "", "", "", out errorText);
                    if (taxDataTable.Rows.Count == 0)
                    {
                        //errorText = "ანგარიშ-ფაქტურა ამ პერიოდში არ მოიძებნა!"; //თუ რს-ზე იძებნება ა/ფ, მაშინ რადგან სტატუსი აქვს უარყოფილი რს-ზე მაგიტომ ვერ პოულობს ეს სერვისი (სავარაუდოდ).
                        return;
                    }
                    taxDataRow = taxDataTable.Rows[0];
                }

                object ID = taxDataRow["ID"]; // ანგარიშ-ფაქტურის უნიკალური ნომერი
                object BUYER_UN_ID = taxDataRow["BUYER_UN_ID"]; // მყიდველის გადამხდელის უნიკალური ნომერი
                string SEQ_NUM_S = taxDataRow["SEQ_NUM_S"].ToString(); // მყიდველის დეკლარაციის ნომერი
                object STATUS = taxDataRow["STATUS"]; // ანგარიშ-ფაქტურის სტატუსი
                int WAS_REF = Convert.ToInt32(string.IsNullOrEmpty(taxDataRow["WAS_REF"].ToString()) ? 0 : taxDataRow["WAS_REF"]); // უარყოფილი მეორე მხარის მიერ 0 - არა 1 - კი
                object F_SERIES = taxDataRow["F_SERIES"]; // ანგარიშ-ფაქტურის სერია
                object F_NUMBER = taxDataRow["F_NUMBER"]; // ანგარიშ-ფაქტურის ნომერი
                DateTime REG_DT = new DateTime();
                REG_DT = DateTime.TryParse(taxDataRow["REG_DT"].ToString(), out REG_DT) == false ? new DateTime() : REG_DT; // რეგისტრაციის თარიღი
                DateTime OPERATION_DT = new DateTime();
                OPERATION_DT = DateTime.TryParse(taxDataRow["OPERATION_DT"].ToString(), out OPERATION_DT) == false ? new DateTime() : OPERATION_DT; // ოპერაციის განხორციელების თარიღი
                object S_USER_ID = taxDataRow["S_USER_ID"]; // სერვისის მომხმარებლის უნიკალური ნომერი
                //object B_S_USER_ID = taxDataRow["B_S_USER_ID"]; // მყიდველის სერვისის მომხმარებლის უნიკალური ნომერი
                object DOC_MOS_NOM_S = taxDataRow["DOC_MOS_NOM_S"]; // ??? 
                object SA_IDENT_NO = taxDataRow["SA_IDENT_NO"]; // მყიდველის საიდენტიფიკაციო ნომერი
                object ORG_NAME = taxDataRow["ORG_NAME"]; // მყიდველის დასახელება 
                object NOTES = taxDataRow["NOTES"]; // მყიდველის მაღაზიის ნომერი
                double TANXA = Convert.ToDouble(string.IsNullOrEmpty(taxDataRow["TANXA"].ToString()) ? 0 : taxDataRow["TANXA"], CultureInfo.InvariantCulture); // თანხა  დღგ-ის ჩათვლით
                double VAT = Convert.ToDouble(string.IsNullOrEmpty(taxDataRow["VAT"].ToString()) ? 0 : taxDataRow["VAT"], CultureInfo.InvariantCulture); // დღგ-ის თანხა
                string K_ID = taxDataRow["K_ID"].ToString(); // კორექტირების ანგარიშ-ფაქტურის ID
                DateTime AGREE_DATE = new DateTime();
                AGREE_DATE = DateTime.TryParse(taxDataRow["AGREE_DATE"].ToString(), out AGREE_DATE) == false ? new DateTime() : AGREE_DATE; // დადასტურების თარიღი
                object AGREE_S_USER_ID = taxDataRow["AGREE_S_USER_ID"]; // დამდასტურებელი
                DateTime REF_DATE = new DateTime();
                REF_DATE = DateTime.TryParse(taxDataRow["REF_DATE"].ToString(), out REF_DATE) == false ? new DateTime() : REF_DATE; // უარყოფის თარიღი
                object REF_S_USER_ID = taxDataRow["REF_S_USER_ID"]; // უარმყოფელი

                if (STATUS != null)
                {
                    bool corrInv = (K_ID != "-1"); //თუ არის კორექტირების ა/ფ
                    bool refInv = (WAS_REF == 1); //თუ არის უარყოფილი ა/ფ

                    STATUS = getStatusValueByStatusNumber(STATUS.ToString(), corrInv, refInv);

                    oGeneralData.SetProperty("U_status", STATUS);
                    oGeneralData.SetProperty("U_declNumber", SEQ_NUM_S);
                    //oGeneralData.SetProperty("U_vatRDate", declDate);
                    oGeneralData.SetProperty("U_number", F_NUMBER.ToString());
                    oGeneralData.SetProperty("U_series", F_SERIES.ToString());
                    oGeneralData.SetProperty("U_sentDate", REG_DT);
                    oGeneralData.SetProperty("U_opDate", OPERATION_DT);
                    oGeneralData.SetProperty("U_confDate", AGREE_DATE);
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

        /// <summary>განახლება</summary>
        private static void update_invoice(TaxInvoice oTaxInvoice, SAPbobsCOM.GeneralData oGeneralData, SAPbobsCOM.GeneralData oGeneralDataCorr, string operation, out string errorText, out string errorTextWb, out string errorTextGoods)
        {
            errorText = null;
            errorTextWb = null;
            errorTextGoods = null;

            try
            {
                string invID = oGeneralData.GetProperty("U_invID");
                bool corrInv = oGeneralData.GetProperty("U_corrInv") == "N" ? false : true;
                string status = oGeneralData.GetProperty("U_status");

                if (string.IsNullOrEmpty(invID))
                {
                    errorText = BDOSResources.getTranslate("TaxInvoiceIDNotFilled");
                    return;
                }

                int inv_ID = Convert.ToInt32(invID);

                //ზედნადებების ცხრილის დამატება
                bool response = save_ntos_invoices_inv_nos(oTaxInvoice, oGeneralData, oGeneralDataCorr, inv_ID, out errorTextWb);
                if (response == false)
                {
                    status = "9"; //არასრულად შექმნილი
                }

                //საქონლის ცხრილის დამატება
                response = save_invoice_desc(oTaxInvoice, oGeneralData, inv_ID, out errorTextGoods);
                if (response == false)
                {
                    status = "9"; //არასრულად შექმნილი
                }
                if (status == "9") //არასრულად შექმნილი
                {
                    bool refInv = false;
                    status = getStatusValueByStatusNumber(status, corrInv, refInv);
                    oGeneralData.SetProperty("U_status", status);
                }
                else
                {
                    get_invoice(oTaxInvoice, oGeneralData, operation, out errorText);
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
