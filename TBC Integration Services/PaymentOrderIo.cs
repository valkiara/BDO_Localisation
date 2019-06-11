using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn.TBC_Integration_Services
{
    /// <remarks/>
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(TransferToOtherBankForeignCurrencyPaymentOrderIo))]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(TransferToOtherBankNationalCurrencyPaymentOrderIo))]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(TreasuryTransferPaymentOrderIo))]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(TransferWithinBankPaymentOrderIo))]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(TransferToOwnAccountPaymentOrderIo))]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://www.mygemini.com/schemas/mygemini")]
    public abstract partial class PaymentOrderIo : AbstractIo
    {

        private AccountIdentificationIo creditAccountField;

        private AccountIdentificationIo debitAccountField;

        private long documentNumberField;

        private bool documentNumberFieldSpecified;

        private MoneyIo amountField;

        private int positionField;

        private string additionalDescriptionField;

        private string descriptionField;

        /// <remarks/>
        public AccountIdentificationIo creditAccount
        {
            get
            {
                return this.creditAccountField;
            }
            set
            {
                this.creditAccountField = value;
            }
        }

        /// <remarks/>
        public AccountIdentificationIo debitAccount
        {
            get
            {
                return this.debitAccountField;
            }
            set
            {
                this.debitAccountField = value;
            }
        }

        /// <remarks/>
        public long documentNumber
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
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool documentNumberSpecified
        {
            get
            {
                return this.documentNumberFieldSpecified;
            }
            set
            {
                this.documentNumberFieldSpecified = value;
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
        public int position
        {
            get
            {
                return this.positionField;
            }
            set
            {
                this.positionField = value;
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
    }
}
