using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn.TBC_Integration_Services
{
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://www.mygemini.com/schemas/mygemini")]
    public partial class AccountMovementDetailIo : AbstractIo
    {

        private string movementIdField;

        private string paymentIdField;

        private string externalPaymentIdField;

        private int debitCreditField;

        private System.DateTime valueDateField;

        private string descriptionField;

        private MoneyIo amountField;

        private string accountNumberField;

        private string accountNameField;

        private string additionalInformationField;

        private System.DateTime documentDateField;

        private bool documentDateFieldSpecified;

        private string documentNumberField;

        private string partnerAccountNumberField;

        private string partnerNameField;

        private string partnerTaxCodeField;

        private string partnerBankCodeField;

        private string partnerBankField;

        private string intermediaryBankCodeField;

        private string intermediaryBankField;

        private string chargeDetailField;

        private string taxpayerCodeField;

        private string taxpayerNameField;

        private string treasuryCodeField;

        private string operationCodeField;

        private string additionalDescriptionField;

        private string exchangeRateField;

        private string partnerPersonalNumberField;

        private string partnerDocumentTypeField;

        private string partnerDocumentNumberField;

        private string parentExternalPaymentIdField;

        private string statusCodeField;

        private string transactionTypeField;

        /// <remarks/>
        public string movementId
        {
            get
            {
                return this.movementIdField;
            }
            set
            {
                this.movementIdField = value;
            }
        }

        /// <remarks/>
        public string paymentId
        {
            get
            {
                return this.paymentIdField;
            }
            set
            {
                this.paymentIdField = value;
            }
        }

        /// <remarks/>
        public string externalPaymentId
        {
            get
            {
                return this.externalPaymentIdField;
            }
            set
            {
                this.externalPaymentIdField = value;
            }
        }

        /// <remarks/>
        public int debitCredit
        {
            get
            {
                return this.debitCreditField;
            }
            set
            {
                this.debitCreditField = value;
            }
        }

        /// <remarks/>
        public System.DateTime valueDate
        {
            get
            {
                return this.valueDateField;
            }
            set
            {
                this.valueDateField = value;
            }
        }

        /// <remarks/>
        public string description
        {
            get
            {
                return this.descriptionField;
            }
            set
            {
                this.descriptionField = value;
            }
        }

        /// <remarks/>
        public MoneyIo amount
        {
            get
            {
                return this.amountField;
            }
            set
            {
                this.amountField = value;
            }
        }

        /// <remarks/>
        public string accountNumber
        {
            get
            {
                return this.accountNumberField;
            }
            set
            {
                this.accountNumberField = value;
            }
        }

        /// <remarks/>
        public string accountName
        {
            get
            {
                return this.accountNameField;
            }
            set
            {
                this.accountNameField = value;
            }
        }

        /// <remarks/>
        public string additionalInformation
        {
            get
            {
                return this.additionalInformationField;
            }
            set
            {
                this.additionalInformationField = value;
            }
        }

        /// <remarks/>
        public System.DateTime documentDate
        {
            get
            {
                return this.documentDateField;
            }
            set
            {
                this.documentDateField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool documentDateSpecified
        {
            get
            {
                return this.documentDateFieldSpecified;
            }
            set
            {
                this.documentDateFieldSpecified = value;
            }
        }

        /// <remarks/>
        public string documentNumber
        {
            get
            {
                return this.documentNumberField;
            }
            set
            {
                this.documentNumberField = value;
            }
        }

        /// <remarks/>
        public string partnerAccountNumber
        {
            get
            {
                return this.partnerAccountNumberField;
            }
            set
            {
                this.partnerAccountNumberField = value;
            }
        }

        /// <remarks/>
        public string partnerName
        {
            get
            {
                return this.partnerNameField;
            }
            set
            {
                this.partnerNameField = value;
            }
        }

        /// <remarks/>
        public string partnerTaxCode
        {
            get
            {
                return this.partnerTaxCodeField;
            }
            set
            {
                this.partnerTaxCodeField = value;
            }
        }

        /// <remarks/>
        public string partnerBankCode
        {
            get
            {
                return this.partnerBankCodeField;
            }
            set
            {
                this.partnerBankCodeField = value;
            }
        }

        /// <remarks/>
        public string partnerBank
        {
            get
            {
                return this.partnerBankField;
            }
            set
            {
                this.partnerBankField = value;
            }
        }

        /// <remarks/>
        public string intermediaryBankCode
        {
            get
            {
                return this.intermediaryBankCodeField;
            }
            set
            {
                this.intermediaryBankCodeField = value;
            }
        }

        /// <remarks/>
        public string intermediaryBank
        {
            get
            {
                return this.intermediaryBankField;
            }
            set
            {
                this.intermediaryBankField = value;
            }
        }

        /// <remarks/>
        public string chargeDetail
        {
            get
            {
                return this.chargeDetailField;
            }
            set
            {
                this.chargeDetailField = value;
            }
        }

        /// <remarks/>
        public string taxpayerCode
        {
            get
            {
                return this.taxpayerCodeField;
            }
            set
            {
                this.taxpayerCodeField = value;
            }
        }

        /// <remarks/>
        public string taxpayerName
        {
            get
            {
                return this.taxpayerNameField;
            }
            set
            {
                this.taxpayerNameField = value;
            }
        }

        /// <remarks/>
        public string treasuryCode
        {
            get
            {
                return this.treasuryCodeField;
            }
            set
            {
                this.treasuryCodeField = value;
            }
        }

        /// <remarks/>
        public string operationCode
        {
            get
            {
                return this.operationCodeField;
            }
            set
            {
                this.operationCodeField = value;
            }
        }

        /// <remarks/>
        public string additionalDescription
        {
            get
            {
                return this.additionalDescriptionField;
            }
            set
            {
                this.additionalDescriptionField = value;
            }
        }

        /// <remarks/>
        public string exchangeRate
        {
            get
            {
                return this.exchangeRateField;
            }
            set
            {
                this.exchangeRateField = value;
            }
        }

        /// <remarks/>
        public string partnerPersonalNumber
        {
            get
            {
                return this.partnerPersonalNumberField;
            }
            set
            {
                this.partnerPersonalNumberField = value;
            }
        }

        /// <remarks/>
        public string partnerDocumentType
        {
            get
            {
                return this.partnerDocumentTypeField;
            }
            set
            {
                this.partnerDocumentTypeField = value;
            }
        }

        /// <remarks/>
        public string partnerDocumentNumber
        {
            get
            {
                return this.partnerDocumentNumberField;
            }
            set
            {
                this.partnerDocumentNumberField = value;
            }
        }

        /// <remarks/>
        public string parentExternalPaymentId
        {
            get
            {
                return this.parentExternalPaymentIdField;
            }
            set
            {
                this.parentExternalPaymentIdField = value;
            }
        }

        /// <remarks/>
        public string statusCode
        {
            get
            {
                return this.statusCodeField;
            }
            set
            {
                this.statusCodeField = value;
            }
        }

        /// <remarks/>
        public string transactionType
        {
            get
            {
                return this.transactionTypeField;
            }
            set
            {
                this.transactionTypeField = value;
            }
        }
    }

}
