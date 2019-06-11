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
    public partial class GetPaymentOrderStatusResponseIo
    {

        private string statusField;

        private PaymentStatusDataIo singlePaymentDataField;

        private PaymentStatusDataIo[] batchPaymentDataField;

        /// <remarks/>
        public string status
        {
            get
            {
                return this.statusField;
            }
            set
            {
                this.statusField = value;
            }
        }

        /// <remarks/>
        public PaymentStatusDataIo singlePaymentData
        {
            get
            {
                return this.singlePaymentDataField;
            }
            set
            {
                this.singlePaymentDataField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("batchPaymentData")]
        public PaymentStatusDataIo[] batchPaymentData
        {
            get
            {
                return this.batchPaymentDataField;
            }
            set
            {
                this.batchPaymentDataField = value;
            }
        }
    }

}
