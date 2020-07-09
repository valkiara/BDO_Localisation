using BDO_Localisation_AddOn.BOG_Integration_Services.Model;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

//using Newtonsoft.Json.Linq;

namespace BDO_Localisation_AddOn.BOG_Integration_Services
{
    static partial class MainPaymentServiceBOG
    {
        /// <summary>ეროვნულ ვალუტაში ინდივიდუალური დოკუმენტის შექმნა</summary>
        /// <param name="oPaymentOrderIo"></param>
        /// <param name="dataForImport"></param>
        /// <param name="ImportBatchPaymentOrders"></param>
        public static void createDomesticPaymentOrderIo(DomesticPayment oPaymentOrderIo, Dictionary<string, object> dataForImport, bool ImportBatchPaymentOrders = false)
        {
            CultureInfo culture = new CultureInfo("en-US");
            string transferType = dataForImport["TransferType"].ToString();

            oPaymentOrderIo.SourceAccountNumber = dataForImport["DebitAccount"].ToString(); //გამგზავნის ანგარიში, ანგარიშიდან / მხოლოდ IBAN-ის ფორმატი
            oPaymentOrderIo.Amount = Convert.ToDecimal(dataForImport["Amount"], culture); //თანხა
            oPaymentOrderIo.BeneficiaryAccountNumber = dataForImport["CreditAccount"].ToString(); //მიმღების ანგარიში, ანგარიშზე ეროვნული ვალუტის გადარიცხვისას საქართველოს ბანკში, სხვა ბანკში – მხოლოდ IBAN-ის ფორმატი. ხაზინაში გადარიცხვისას მიუთითეთ სახაზინო კოდი უცხოური ვალუტის გადარიცხვისას მიმრების ანგარიშის ნომერი არსებული ფორმატით
            oPaymentOrderIo.BeneficiaryBankCode = dataForImport["BeneficiaryBankCode"].ToString(); //მიმღები ბანკის RTGS კოდი / სავალდებულო            
            oPaymentOrderIo.BeneficiaryName = dataForImport["BeneficiaryName"].ToString(); //მიმღების დასახელება
            oPaymentOrderIo.BeneficiaryInn = dataForImport["BeneficiaryTaxCode"].ToString(); //მიმღების – ფიზიკური პირის პირადი ნომერი, იურიდიული პირის საიდენტიფიკაციო კოდი          
            oPaymentOrderIo.DispatchType = dataForImport["DispatchType"].ToString(); //გადარიცხვის მეთოდი.BULK - სტანდარტული გადარიცხვა, MT103 ინდივიდუალური გადარიცხვა. 10 000 ლარამდე შესაძლებელია გამოიყენოთ ორივე მეთოდი. 10 000 ზემოთ მხოლოდ MT103
            oPaymentOrderIo.DocumentNo = dataForImport["DocEntry"].ToString(); //დოკ. N
            oPaymentOrderIo.Nomination = dataForImport["Description"].ToString(); //გადარიცხვის საფუძველი, დანიშნულება / მაქსიმუმ 250 სიმბოლო
            oPaymentOrderIo.AdditionalInformation = ""; //დამატებითი ინფორმაცია / არა სავალდებულო / არ გამოიყენება ეროვნული ვალუტის გადარიცხვისას
            oPaymentOrderIo.PayerInn = dataForImport["TaxpayerCode"].ToString(); //გადამხდელის – ფიზიკური პირის პირადი ნომერი, იურიდიული პირის საიდენტიფიკაციო კოდი ივსება მხოლოდ! ბიუჯეტის სასარგებლოდ მესამე პირის ნაცვლად/მაგივრად გადარიცხვის შესრულებისას, მიუთითეთ მესამე პირის საიდენტიფიკაციო კოდი
            oPaymentOrderIo.PayerName = dataForImport["TaxpayerName"].ToString(); //გადამხდელის დასახელება ივსება მხოლოდ! ბიუჯეტის სასარგებლოდ მესამე პირის ნაცვლად/მაგივრად გადარიცხვის შესრულებისას, მიუთითეთ მესამე პირის დასახელება
            oPaymentOrderIo.ValueDate = DateTime.Today; //ვალუტირების თარიღი
            //oPaymentOrderIo.IsSalary = false;//არასავალდებულო ველი, სასურველია შეივსოს სახელფასო გადარიცხვისას ხელფასის გადარიცხვისას "TRUE" სხვა შემთხვევაში "FALSE"
            oPaymentOrderIo.UniqueId = Guid.NewGuid(); //უნიკალური იდენტიფიკატორი (Guid) გარე სისტემაში
            //oPaymentOrderIo.CheckInn = false; //“საქართველოს ბანკში“ გადარიცხვის შემოწმება საიდენტიფიკაციო კოდით არასავალდებულო ველი “საქართველოს ბანკში“ გადარიცხვისას “TRUE” მითითების შემთხვევაში, დამატებით შემოწმდება მიმღების ანგარიშის ნომრი მითითებული საიდენტიფიკაციო კოდით, თუ არ ანგარიშის ნომერი არ ეკუთვნის მითითებულ საიდენტიფიკაციო კოდზე არსებულ ანგარიშს დაფიქსირდება შეცდომა. “საქართველოს ბანკში“ გადარიცხვისას “FALSE” მითითების შემთხვევაში, გადარიცხვა შემოწმდება სტანდარტული წესით სხვა ბანკში გადარიცხვისას ამ ველის მნიშვნელობა დადებითი არ უნდა იყოს  
        }

        /// <summary>უცხოურ ვალუტაში ინდივიდუალური დოკუმენტის შექმნა</summary>
        /// <param name="oPaymentOrderIo"></param>
        /// <param name="dataForImport"></param>
        /// <param name="ImportBatchPaymentOrders"></param>
        public static void createForeignPaymentOrderIo( ForeignPayment oPaymentOrderIo, Dictionary<string, object> dataForImport, bool ImportBatchPaymentOrders = false)
        {
            CultureInfo culture = new CultureInfo("en-US");
            string DebitBankCode = dataForImport["DebitBankCode"].ToString();
            string BeneficiaryBankCode = dataForImport["BeneficiaryBankCode"].ToString();
            string BeneficiaryRegistrationCountryCode = dataForImport["BeneficiaryRegistrationCountryCode"].ToString();
            
            if (BeneficiaryBankCode == DebitBankCode && string.IsNullOrEmpty(BeneficiaryRegistrationCountryCode))
            {
                BeneficiaryRegistrationCountryCode = CommonFunctions.getRegistrationCountryCode( dataForImport["CreditAccount"].ToString() + dataForImport["CreditAccountCurrencyCode"].ToString(), "OCRB");
            }
            //else 
            //{
            //    BeneficiaryRegistrationCountryCode = CommonFunctions.getRegistrationCountryCode( dataForImport["CreditAccount"].ToString() + dataForImport["CreditAccountCurrencyCode"].ToString(), "DSC1");
            //}
            
            oPaymentOrderIo.SourceAccountNumber = dataForImport["DebitAccount"].ToString(); //გამგზავნის ანგარიში, ანგარიშიდან / მხოლოდ IBAN-ის ფორმატი
            oPaymentOrderIo.Amount = Convert.ToDecimal(dataForImport["Amount"], culture); //თანხა
            oPaymentOrderIo.Currency = dataForImport["Currency"].ToString(); //ვალუტა
            oPaymentOrderIo.BeneficiaryAccountNumber = dataForImport["CreditAccount"].ToString(); //მიმღების ანგარიში, ანგარიშზე ეროვნული ვალუტის გადარიცხვისას საქართველოს ბანკში, სხვა ბანკში – მხოლოდ IBAN-ის ფორმატი. ხაზინაში გადარიცხვისას მიუთითეთ სახაზინო კოდი უცხოური ვალუტის გადარიცხვისას მიმრების ანგარიშის ნომერი არსებული ფორმატით
            oPaymentOrderIo.BeneficiaryBankName = dataForImport["BeneficiaryBankName"].ToString(); //მიმღები ბანკის დასახელება. სავალდებულოა ან BeneficiaryBankCode ან BeneficiaryBankName
            oPaymentOrderIo.BeneficiaryBankCode = BeneficiaryBankCode; //მიმღები ბანკის RTGS კოდი / სავალდებულო
            oPaymentOrderIo.BeneficiaryInn = dataForImport["BeneficiaryTaxCode"].ToString(); //მიმღების – ფიზიკური პირის პირადი ნომერი, იურიდიული პირის საიდენტიფიკაციო კოდი
            oPaymentOrderIo.BeneficiaryName = dataForImport["BeneficiaryName"].ToString(); //მიმღების დასახელება
            oPaymentOrderIo.DocumentNo = dataForImport["DocEntry"].ToString(); //დოკ. N
            oPaymentOrderIo.PaymentDetail = dataForImport["Description"].ToString(); //გადარიცხვის საფუძველი, დანიშნულება / მაქსიმუმ 4x34 136 სიმბოლო
            oPaymentOrderIo.AdditionalInformation = dataForImport["AdditionalDescription"].ToString(); //დამატებითი ინფორმაცია / არა სავალდებულო / არ გამოიყენება ეროვნული ვალუტის გადარიცხვისას
            oPaymentOrderIo.ValueDate = DateTime.Today; //ვალუტირების თარიღი
            oPaymentOrderIo.UniqueId = Guid.NewGuid(); //უნიკალური იდენტიფიკატორი (Guid) გარე სისტემაში
            //oPaymentOrderIo.CheckInn = false; //“საქართველოს ბანკში“ გადარიცხვის შემოწმება საიდენტიფიკაციო კოდით არასავალდებულო ველი “საქართველოს ბანკში“ გადარიცხვისას “TRUE” მითითების შემთხვევაში, დამატებით შემოწმდება მიმღების ანგარიშის ნომრი მითითებული საიდენტიფიკაციო კოდით, თუ არ ანგარიშის ნომერი არ ეკუთვნის მითითებულ საიდენტიფიკაციო კოდზე არსებულ ანგარიშს დაფიქსირდება შეცდომა. “საქართველოს ბანკში“ გადარიცხვისას “FALSE” მითითების შემთხვევაში, გადარიცხვა შემოწმდება სტანდარტული წესით სხვა ბანკში გადარიცხვისას ამ ველის მნიშვნელობა დადებითი არ უნდა იყოს  
            oPaymentOrderIo.BeneficiaryActualCountryCode = BeneficiaryRegistrationCountryCode; //api.property.foreign.benef.actual.country
            oPaymentOrderIo.BeneficiaryRegistrationCountryCode = BeneficiaryRegistrationCountryCode; //api.property.foreign.benef.reg.country
            oPaymentOrderIo.IntermediaryBankName = dataForImport["IntermediaryBankName"].ToString(); //შუამავალი ბანკის დასახელება
            oPaymentOrderIo.IntermediaryBankCode = dataForImport["IntermediaryBankCode"].ToString(); //შუამავალი ბანკის SWIFT კოდი
            if (!ImportBatchPaymentOrders)
            {
                oPaymentOrderIo.ComissionAccountNumber = ""; //საკომისიოს ანგარიში/მხოლოდ ინდივიდუალური გადარიცხვისთვის
            }
            oPaymentOrderIo.Charges = dataForImport["ChargeDetails"].ToString(); //გადარიცხვის მეთოდი (SHA ან OUR)
            oPaymentOrderIo.RegReportingValue = dataForImport["reportCode"].ToString(); //reporting code (GDS, ACM, DCM, AKA)

            oPaymentOrderIo.RecipientCity = dataForImport["RecipientCity"].ToString();
            oPaymentOrderIo.RecipientAddress = dataForImport["BeneficiaryAddress"].ToString();
        }

        /// <summary>კონვერტაცია</summary>
        /// <param name="oPaymentOrderIo"></param>
        /// <param name="dataForImport"></param>
        /// <param name="ImportBatchPaymentOrders"></param>
        public static void createConversionPaymentOrderIo(ConversionPayment oPaymentOrderIo, Dictionary<string, object> dataForImport)
        {
            CultureInfo culture = new CultureInfo("en-US");

            oPaymentOrderIo.SourceAccountNumber = dataForImport["DebitAccount"].ToString(); //გაყიდვის ანგარიში, ანგარიშიდან
            oPaymentOrderIo.SourceCurrency = dataForImport["DebitAccountCurrencyCode"].ToString(); //გასაყიდი ვალუტა
            oPaymentOrderIo.Amount = Convert.ToDecimal(dataForImport["Amount"], culture); // გასაყიდი თანხა
            oPaymentOrderIo.DestinationAccountNumber = dataForImport["CreditAccount"].ToString(); //ყიდვის ანგარიში, ყიდვა ანგარიშზე
            oPaymentOrderIo.DestinationCurrency = dataForImport["CreditAccountCurrencyCode"].ToString(); //საყიდელი ვალუტა, ყიდვის ვალუტა
            oPaymentOrderIo.Rate = Convert.ToDecimal(dataForImport["DocRate"], culture); //კონვერტაციის კურსი
            oPaymentOrderIo.DocumentNo = dataForImport["DocEntry"].ToString(); //დოკ. N
            oPaymentOrderIo.UniqueId = Guid.NewGuid(); //უნიკალური იდენტიფიკატორი (Guid) გარე სისტემაში
            oPaymentOrderIo.AdditionalInfo = dataForImport["AdditionalDescription"].ToString(); //დამატებითი ინფორმაცია / არა სავალდებულო / არ გამოიყენება ეროვნული ვალუტის გადარიცხვისას
        }

        public async static Task<DocumentKey[]> importConversionPaymentOrders(HttpClient client, List<ConversionPayment> paymentOrderList)
        {
            string errorText = null;
            DocumentKey[] keys = null;

            //XML --->
            //string text = @"<ArrayOfConversionPayment xmlns:i=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://schemas.datacontract.org/2004/07/CIBApi.Models"">
            //<ConversionPayment>
            //<AdditionalInfo>salome</AdditionalInfo>
            //<Amount>7</Amount>
            //<DestinationAccountNumber>GE91BG0000000853540002</DestinationAccountNumber>
            //<DestinationCurrency>EUR</DestinationCurrency>
            //<DocumentNo>12345</DocumentNo>
            //<Rate>2.6516</Rate>
            //<SourceAccountNumber>GE91BG0000000853540002</SourceAccountNumber>
            //<SourceCurrency>GEL</SourceCurrency>
            //<UniqueId>cbd462fc-89dd-4d2f-a11b-8b0a907e7736</UniqueId>
            //</ConversionPayment>
            //</ArrayOfConversionPayment>";
            //var httpContent = new StringContent(text, Encoding.UTF8, "application/xml");
            //var respone = await client.PostAsync("documents/conversion", httpContent);
            //if (!respone.IsSuccessStatusCode)
            //{
            ////throw new InvalidUriException(string.Format("Invalid uri: {0}", requestUri));
            //}
            //XML --->

            var response = await client.PostAsJsonAsync("documents/conversion", paymentOrderList);
            if (response.IsSuccessStatusCode)
            {
                keys = await response.Content.ReadAsAsync<DocumentKey[]>();
            }
            else
            {
                errorText = await ShowErrorMessage(response);
            }

            return keys;
        }

        public async static Task<DocumentKey[]> importForeignPaymentOrders(HttpClient client, List<ForeignPayment> paymentOrderList)
        {
            string errorText = null;
            DocumentKey[] keys = null;

            var response = await client.PostAsJsonAsync("documents/foreign", paymentOrderList);
            if (response.IsSuccessStatusCode)
            {
                keys = await response.Content.ReadAsAsync<DocumentKey[]>();
            }
            else
            {
                errorText = await ShowErrorMessage(response);
            }

            return keys;
        }

        public async static Task<long> importBulkForeignPaymentOrders(HttpClient client, List<ForeignPayment> paymentOrderList)
        {
            string errorText = null;
            long key = 0;

            var response = await client.PostAsJsonAsync("documents/bulk/foreign", paymentOrderList);
            if (response.IsSuccessStatusCode)
            {
                key = await response.Content.ReadAsAsync<long>();
            }
            else
            {
                errorText = await ShowErrorMessage(response);
            }

            return key;
        }

        public async static Task<DocumentKey[]> importDomesticPaymentOrders(HttpClient client, List<DomesticPayment> paymentOrderList)
        {
            string errorText = null;
            DocumentKey[] keys = null;

            var response = await client.PostAsJsonAsync("documents/domestic", paymentOrderList);
            if (response.IsSuccessStatusCode)
            {
                keys = await response.Content.ReadAsAsync<DocumentKey[]>();
            }
            else
            {
                errorText = await ShowErrorMessage(response);
            }

            return keys;
        }

        public async static Task<long> importBulkDomesticPaymentOrders(HttpClient client, List<DomesticPayment> paymentOrderList)
        {
            string errorText = null;
            long key = 0;

            var response = await client.PostAsJsonAsync("documents/bulk/domestic", paymentOrderList);
            if (response.IsSuccessStatusCode)
            {
                key = await response.Content.ReadAsAsync<long>();
            }
            else
            {
                errorText = await ShowErrorMessage(response);
            }

            return key;
        }

        public async static Task<List<DocumentStatus>> refreshSinglePaymentOrderStatus(HttpClient client, long key)
        {
            string errorText = null;
            List<DocumentStatus> oDocumentStatus = null;

            var response = await client.GetAsync(string.Format("documents/statuses/{0}", key));

            if (response.IsSuccessStatusCode)
            {
                oDocumentStatus = await response.Content.ReadAsAsync<List<DocumentStatus>>();               
            }
            else
            {
                errorText = await ShowErrorMessage(response);
            }

            return oDocumentStatus;
        }

        public async static Task<BulkPaymentStatus> refreshBatchPaymentOrderStatus(HttpClient client, string bulkID)
        {
            string errorText = null;
            BulkPaymentStatus oBulkPaymentStatus = null;

            var response = await client.GetAsync(string.Format("documents/bulk/status/{0}", bulkID));

            if (response.IsSuccessStatusCode)
            {
                oBulkPaymentStatus = await response.Content.ReadAsAsync<BulkPaymentStatus>();
            }
            else
            {
                errorText = await ShowErrorMessage(response);
            }

            return oBulkPaymentStatus;
        }

        public async static Task<StatementSummary> getStatementSummary(HttpClient client, string accountNumber, string currency, DateTime periodFrom, DateTime periodTo)
        {
            string errorText = null;
            StatementSummary summary = null;

            string format = "yyyy-MM-dd";

            var response = await client.GetAsync(string.Format("statement/{0}/{1}/{2}/{3}", accountNumber, currency,
                                                 periodFrom.ToString(format), periodTo.ToString(format)));
            if (response.IsSuccessStatusCode)
            {
                var statement = await response.Content.ReadAsAsync<Statement>();
                var summaryResponse = await client.GetAsync(String.Format("statement/summary/{0}/{1}/{2}/", accountNumber, currency, statement.Id));

                if (summaryResponse.IsSuccessStatusCode)
                {
                    summary = await summaryResponse.Content.ReadAsAsync<StatementSummary>();
                }
            }
            else
            {
                errorText = await ShowErrorMessage(response);
            }
            return summary;
        }

        public async static Task<List<StatementDetail>> getStatement(HttpClient client, string accountNumber, string currency, DateTime periodFrom, DateTime periodTo, int page)
        {
            string errorText = null;
            List<StatementDetail> summary = null;

            string format = "yyyy-MM-dd";

            var response = await client.GetAsync(string.Format("statement/{0}/{1}/{2}/{3}", accountNumber, currency,
                                                 periodFrom.ToString(format), periodTo.ToString(format)));
            if (response.IsSuccessStatusCode)
            {
                var statement = await response.Content.ReadAsAsync<Statement>();

                var summaryResponse = await client.GetAsync(String.Format("statement/{0}/{1}/{2}/{3}/", accountNumber, currency, statement.Id, page));

                if (summaryResponse.IsSuccessStatusCode)
                {
                    summary = await summaryResponse.Content.ReadAsAsync<List<StatementDetail>>();
                }
            }
            else
            {
                errorText = await ShowErrorMessage(response);
            }
            return summary;
        }

        public async static Task<Statement> getStatement(HttpClient client, string accountNumber, string currency, DateTime periodFrom, DateTime periodTo)
        {
            Statement statement = null;

            string format = "yyyy-MM-dd";

            var response = await client.GetAsync(string.Format("statement/{0}/{1}/{2}/{3}", accountNumber, currency,
                                                 periodFrom.ToString(format), periodTo.ToString(format)));
            if (response.IsSuccessStatusCode)
            {
                statement = await response.Content.ReadAsAsync<Statement>();
            }
            else
            {
                throw new Exception(await ShowErrorMessage(response));
            }
            return statement;
        }

        private static async Task<string> ShowErrorMessage(HttpResponseMessage response)
        {
            var content = await response.Content.ReadAsStringAsync();
            return (string.IsNullOrEmpty(content) ? response.ReasonPhrase : content);
        }
    }
}
