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
    public partial class TransferToOtherBankForeignCurrencyPaymentOrderIo : PaymentOrderIo
    {

        private string beneficiaryNameField;

        private string beneficiaryAddressField;

        private string beneficiaryBankCodeField;

        private string beneficiaryBankNameField;

        private string intermediaryBankCodeField;

        private string intermediaryBankNameField;

        private string chargeDetailsField;

        /// <remarks/>
        public string beneficiaryName
        {
            get
            {
                return this.beneficiaryNameField;
            }
            set
            {
                this.beneficiaryNameField = value;
            }
        }

        /// <remarks/>
        public string beneficiaryAddress
        {
            get
            {
                return this.beneficiaryAddressField;
            }
            set
            {
                this.beneficiaryAddressField = value;
            }
        }

        /// <remarks/>
        public string beneficiaryBankCode
        {
            get
            {
                return this.beneficiaryBankCodeField;
            }
            set
            {
                this.beneficiaryBankCodeField = value;
            }
        }

        /// <remarks/>
        public string beneficiaryBankName
        {
            get
            {
                return this.beneficiaryBankNameField;
            }
            set
            {
                this.beneficiaryBankNameField = value;
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
        public string intermediaryBankName
        {
            get
            {
                return this.intermediaryBankNameField;
            }
            set
            {
                this.intermediaryBankNameField = value;
            }
        }

        /// <remarks/>
        public string chargeDetails
        {
            get
            {
                return this.chargeDetailsField;
            }
            set
            {
                this.chargeDetailsField = value;
            }
        }
    }

}
