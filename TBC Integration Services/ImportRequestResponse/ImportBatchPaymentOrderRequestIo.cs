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
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "http://www.mygemini.com/schemas/mygemini")]
    public partial class ImportBatchPaymentOrderRequestIo : AbstractIo
    {

        private AccountIdentificationIo debitAccountIdentificationField;

        private string batchNameField;

        private PaymentOrderIo[] paymentOrderField;

        /// <remarks/>
        public AccountIdentificationIo debitAccountIdentification
        {
            get
            {
                return this.debitAccountIdentificationField;
            }
            set
            {
                this.debitAccountIdentificationField = value;
            }
        }

        /// <remarks/>
        public string batchName
        {
            get
            {
                return this.batchNameField;
            }
            set
            {
                this.batchNameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("paymentOrder")]
        public PaymentOrderIo[] paymentOrder
        {
            get
            {
                return this.paymentOrderField;
            }
            set
            {
                this.paymentOrderField = value;
            }
        }
    }

}
