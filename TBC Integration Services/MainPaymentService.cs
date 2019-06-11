using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Web.Services.Protocols;

namespace BDO_Localisation_AddOn.TBC_Integration_Services
{
    static partial class MainPaymentService
    {
        ///<summary>სტატუსების განახლება (ინდივიდუალური)</summary>
        /// <param name="oPaymentService"></param>
        /// <param name="singlePaymentId"></param>
        /// <param name="singlePaymentIdSpecified"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public static GetPaymentOrderStatusResponseIo refreshSinglePaymentOrderStatus(PaymentService oPaymentService, long singlePaymentId, bool singlePaymentIdSpecified, out string errorText)
        {
            errorText = null;

            GetPaymentOrderStatusRequestIo orders = new GetPaymentOrderStatusRequestIo();
            orders.singlePaymentId = singlePaymentId;
            orders.singlePaymentIdSpecified = singlePaymentIdSpecified;

            GetPaymentOrderStatusResponseIo orderResult = null;

            try
            {
                orderResult = oPaymentService.GetPaymentOrderStatus(orders);
            }
            catch (Exception ex)
            {
                try
                {
                    errorText = ex.Message + '\n' + ((System.Web.Services.Protocols.SoapException)ex).Code.Name;
                    return orderResult;
                }
                catch
                {
                    errorText = ex.Message;
                    return orderResult;
                }
            }

            return orderResult;
        }

        ///<summary>სტატუსების განახლება (პაკეტური)</summary>
        /// <param name="oPaymentService"></param>
        /// <param name="batchPaymentId"></param>
        /// <param name="batchPaymentIdSpecified"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public static GetPaymentOrderStatusResponseIo refreshBatchPaymentOrderStatus(PaymentService oPaymentService, long batchPaymentId, bool batchPaymentIdSpecified, out string errorText)
        {
            errorText = null;

            GetPaymentOrderStatusRequestIo orders = new GetPaymentOrderStatusRequestIo();
            orders.batchPaymentId = batchPaymentId;
            orders.batchPaymentIdSpecified = batchPaymentIdSpecified;

            GetPaymentOrderStatusResponseIo orderResult = null;

            try
            {
                orderResult = oPaymentService.GetPaymentOrderStatus(orders);            
            }
            catch (Exception ex)
            {
                try
                {
                    errorText = ex.Message + '\n' + ((System.Web.Services.Protocols.SoapException)ex).Code.Name;
                    return orderResult;
                }
                catch
                {
                    errorText = ex.Message;
                    return orderResult;
                }
            }

            return orderResult;
        }

        /// <summary>საგადახდო დავალებების იმპორტი (ინდივიდუალური)</summary>
        /// <param name="oPaymentService"></param>
        /// <param name="singlePaymentOrderArray"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public static ImportSinglePaymentOrdersResponseIo importSinglePaymentOrders(PaymentService oPaymentService, PaymentOrderIo[] singlePaymentOrderArray, out string errorText)
        {
            errorText = null;

            ImportSinglePaymentOrdersRequestIo orders = new ImportSinglePaymentOrdersRequestIo();
            orders.singlePaymentOrder = singlePaymentOrderArray;

            ImportSinglePaymentOrdersResponseIo orderResult = null;

            try
            {
                orderResult = oPaymentService.ImportSinglePaymentOrders(orders);
            }
            catch (Exception ex)
            {
                try
                {
                    errorText = ex.Message + '\n' + ((System.Web.Services.Protocols.SoapException)ex).Code.Name;
                    return orderResult;
                }
                catch
                {
                    errorText = ex.Message;
                    return orderResult;
                }
            }

            return orderResult;
        }

        /// <summary>საგადახდო დავალებების იმპორტი (პაკეტური)</summary>
        /// <param name="oPaymentService"></param>
        /// <param name="batchPaymentOrderArray"></param>
        /// <param name="accountNumber"></param>
        /// <param name="accountCurrencyCode"></param>
        /// <param name="batchName"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public static ImportBatchPaymentOrderResponseIo importBatchPaymentOrders(PaymentService oPaymentService, PaymentOrderIo[] batchPaymentOrderArray, string accountNumber, string accountCurrencyCode, string batchName, out string errorText)
        {
            errorText = null;
           
            ImportBatchPaymentOrderRequestIo orders = new ImportBatchPaymentOrderRequestIo();
            orders.paymentOrder = batchPaymentOrderArray;
            orders.batchName = batchName;

            AccountIdentificationIo oAccountIdentificationIo = new AccountIdentificationIo();
            oAccountIdentificationIo.accountNumber = accountNumber;
            oAccountIdentificationIo.accountCurrencyCode = accountCurrencyCode;
            
            orders.debitAccountIdentification = oAccountIdentificationIo;

            ImportBatchPaymentOrderResponseIo OrderResult = null;
            
            try
            {
                OrderResult = oPaymentService.ImportBatchPaymentOrder(orders);
            }
            catch (Exception ex)
            {
                try
                {
                    errorText = ex.Message + '\n' + ((System.Web.Services.Protocols.SoapException)ex).Code.Name;
                    return OrderResult;
                }
                catch
                {
                    errorText = ex.Message;
                    return OrderResult;
                }
            }

            return OrderResult;
        }

        /// <summary>გადარიცხვა საკუთარ ანგარიშზე/კონვერტაცია</summary>    
        public static void createTransferToOwnAccountPaymentOrderIo(TransferToOwnAccountPaymentOrderIo oPaymentOrderIo, Dictionary<string, object> dataForImport, int position, bool ImportBatchPaymentOrders = false)
        {
            CultureInfo culture = new CultureInfo("en-US");
            if (ImportBatchPaymentOrders == false)
            {
                oPaymentOrderIo.debitAccount = new AccountIdentificationIo();
                oPaymentOrderIo.debitAccount.accountNumber = dataForImport["DebitAccount"].ToString();
                oPaymentOrderIo.debitAccount.accountCurrencyCode = dataForImport["DebitAccountCurrencyCode"].ToString();
            }
            oPaymentOrderIo.creditAccount = new AccountIdentificationIo();
            oPaymentOrderIo.creditAccount.accountNumber = dataForImport["CreditAccount"].ToString();
            oPaymentOrderIo.creditAccount.accountCurrencyCode = dataForImport["CreditAccountCurrencyCode"].ToString();

            oPaymentOrderIo.documentNumber = Convert.ToInt64(dataForImport["DocEntry"]);  //long
            oPaymentOrderIo.documentNumberSpecified = false; //Convert.ToBoolean(oRecordSet.Fields.Item("").Value);
            oPaymentOrderIo.amount = new MoneyIo();
            oPaymentOrderIo.amount.amount = Convert.ToDecimal(dataForImport["Amount"], culture);
            oPaymentOrderIo.amount.currency = dataForImport["Currency"].ToString();

            oPaymentOrderIo.additionalDescription = "";
            oPaymentOrderIo.description = dataForImport["Description"].ToString();
            oPaymentOrderIo.position = position;
        }

        /// <summary>გადარიცხვა თიბისი ბანკის ფილიალებში</summary>
        public static void createTransferWithinBankPaymentOrderIo(TransferWithinBankPaymentOrderIo oPaymentOrderIo, Dictionary<string, object> dataForImport, int position, bool ImportBatchPaymentOrders = false)
        {
            CultureInfo culture = new CultureInfo("en-US");
            if (ImportBatchPaymentOrders == false)
            {
                oPaymentOrderIo.debitAccount = new AccountIdentificationIo();
                oPaymentOrderIo.debitAccount.accountNumber = dataForImport["DebitAccount"].ToString();
                oPaymentOrderIo.debitAccount.accountCurrencyCode = dataForImport["DebitAccountCurrencyCode"].ToString();
            }
            oPaymentOrderIo.creditAccount = new AccountIdentificationIo();
            oPaymentOrderIo.creditAccount.accountNumber = dataForImport["CreditAccount"].ToString();

            oPaymentOrderIo.beneficiaryName = dataForImport["BeneficiaryName"].ToString();

            oPaymentOrderIo.documentNumber = Convert.ToInt64(dataForImport["DocEntry"]);  //long
            oPaymentOrderIo.documentNumberSpecified = false; //Convert.ToBoolean(oRecordSet.Fields.Item("").Value);
            oPaymentOrderIo.amount = new MoneyIo();
            oPaymentOrderIo.amount.amount = Convert.ToDecimal(dataForImport["Amount"], culture);
            oPaymentOrderIo.amount.currency = dataForImport["Currency"].ToString();

            oPaymentOrderIo.additionalDescription = dataForImport["AdditionalDescription"].ToString();
            oPaymentOrderIo.description = dataForImport["Description"].ToString();
            oPaymentOrderIo.position = position;
        }

        /// <summary>გადარიცხვა სხვა ბანკში (ეროვნული ვალუტა)</summary>
        public static void createTransferToOtherBankNationalCurrencyPaymentOrderIo(TransferToOtherBankNationalCurrencyPaymentOrderIo oPaymentOrderIo, Dictionary<string, object> dataForImport, int position, bool ImportBatchPaymentOrders = false)
        {
            CultureInfo culture = new CultureInfo("en-US");
            if (ImportBatchPaymentOrders == false)
            {
                oPaymentOrderIo.debitAccount = new AccountIdentificationIo();
                oPaymentOrderIo.debitAccount.accountNumber = dataForImport["DebitAccount"].ToString();
                oPaymentOrderIo.debitAccount.accountCurrencyCode = dataForImport["DebitAccountCurrencyCode"].ToString();
            }
            oPaymentOrderIo.creditAccount = new AccountIdentificationIo();
            oPaymentOrderIo.creditAccount.accountNumber = dataForImport["CreditAccount"].ToString();

            oPaymentOrderIo.beneficiaryName = dataForImport["BeneficiaryName"].ToString();
            oPaymentOrderIo.beneficiaryTaxCode = dataForImport["BeneficiaryTaxCode"].ToString();

            oPaymentOrderIo.documentNumber = Convert.ToInt64(dataForImport["DocEntry"]);  //long
            oPaymentOrderIo.documentNumberSpecified = false; //Convert.ToBoolean(oRecordSet.Fields.Item("").Value);
            oPaymentOrderIo.amount = new MoneyIo();
            oPaymentOrderIo.amount.amount = Convert.ToDecimal(dataForImport["Amount"], culture);
            oPaymentOrderIo.amount.currency = dataForImport["Currency"].ToString();

            oPaymentOrderIo.additionalDescription = dataForImport["AdditionalDescription"].ToString();
            oPaymentOrderIo.description = dataForImport["Description"].ToString();
            oPaymentOrderIo.position = position;
        }

        /// <summary>გადარიცხვა სხვა ბანკში (უცხოური ვალუტა)</summary>
        public static void createTransferToOtherBankForeignCurrencyPaymentOrderIo(TransferToOtherBankForeignCurrencyPaymentOrderIo oPaymentOrderIo, Dictionary<string, object> dataForImport, int position, bool ImportBatchPaymentOrders = false)
        {
            CultureInfo culture = new CultureInfo("en-US");
            if (ImportBatchPaymentOrders == false)
            {
                oPaymentOrderIo.debitAccount = new AccountIdentificationIo();
                oPaymentOrderIo.debitAccount.accountNumber = dataForImport["DebitAccount"].ToString();
                oPaymentOrderIo.debitAccount.accountCurrencyCode = dataForImport["DebitAccountCurrencyCode"].ToString();
            }
            oPaymentOrderIo.creditAccount = new AccountIdentificationIo();
            oPaymentOrderIo.creditAccount.accountNumber = dataForImport["CreditAccount"].ToString();

            oPaymentOrderIo.beneficiaryName = dataForImport["BeneficiaryName"].ToString();
            oPaymentOrderIo.beneficiaryAddress = dataForImport["BeneficiaryAddress"].ToString();
            oPaymentOrderIo.beneficiaryBankCode = dataForImport["BeneficiaryBankCode"].ToString();
            oPaymentOrderIo.beneficiaryBankName = dataForImport["BeneficiaryBankName"].ToString();

            oPaymentOrderIo.intermediaryBankCode = dataForImport["IntermediaryBankCode"].ToString();
            oPaymentOrderIo.intermediaryBankName = dataForImport["IntermediaryBankName"].ToString();
            oPaymentOrderIo.chargeDetails = dataForImport["ChargeDetails"].ToString();

            oPaymentOrderIo.documentNumber = Convert.ToInt64(dataForImport["DocEntry"]);  //long
            oPaymentOrderIo.documentNumberSpecified = false; //Convert.ToBoolean(oRecordSet.Fields.Item("").Value);
            oPaymentOrderIo.amount = new MoneyIo();
            oPaymentOrderIo.amount.amount = Convert.ToDecimal(dataForImport["Amount"], culture);
            oPaymentOrderIo.amount.currency = dataForImport["Currency"].ToString();

            oPaymentOrderIo.additionalDescription = dataForImport["AdditionalDescription"].ToString();
            oPaymentOrderIo.description = dataForImport["Description"].ToString();
            oPaymentOrderIo.position = position;
        }

        /// <summary>საბიუჯეტო გადარიცხვა</summary>
        public static void createTreasuryTransferPaymentOrderIo(TreasuryTransferPaymentOrderIo oPaymentOrderIo, Dictionary<string, object> dataForImport, int position, bool ImportBatchPaymentOrders = false)
        {
            CultureInfo culture = new CultureInfo("en-US");
            if (ImportBatchPaymentOrders == false)
            {
                oPaymentOrderIo.debitAccount = new AccountIdentificationIo();
                oPaymentOrderIo.debitAccount.accountNumber = dataForImport["DebitAccount"].ToString();
                oPaymentOrderIo.debitAccount.accountCurrencyCode = dataForImport["DebitAccountCurrencyCode"].ToString();
            }

            oPaymentOrderIo.taxpayerCode = dataForImport["TaxpayerCode"].ToString(); //თუ სხვის ნაცვლად იხდი
            oPaymentOrderIo.taxpayerName = dataForImport["TaxpayerName"].ToString(); //თუ სხვის ნაცვლად იხდი

            oPaymentOrderIo.treasuryCode = dataForImport["TreasuryCode"].ToString();

            oPaymentOrderIo.documentNumber = Convert.ToInt64(dataForImport["DocEntry"]);  //long
            oPaymentOrderIo.documentNumberSpecified = false; //Convert.ToBoolean(oRecordSet.Fields.Item("").Value);
            oPaymentOrderIo.amount = new MoneyIo();
            oPaymentOrderIo.amount.amount = Convert.ToDecimal(dataForImport["Amount"], culture);
            oPaymentOrderIo.amount.currency = dataForImport["Currency"].ToString();

            oPaymentOrderIo.additionalDescription = dataForImport["AdditionalDescription"].ToString();
            oPaymentOrderIo.description = "";
            oPaymentOrderIo.position = position;
        }    

        /// <summary>ავტორიზაციის პარამეტრების შევსება</summary>
        /// <param name="serviceUrl"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="nonce"></param>
        /// <returns></returns>
        public static PaymentService setPaymentService(string serviceUrl, string username, string password, string nonce)
        {
            PaymentService oPaymentService = new PaymentService();
            oPaymentService.SetUrl(serviceUrl);
            oPaymentService.SetUsernameToken(username, password, nonce);

            return oPaymentService;
        }
    }
}
