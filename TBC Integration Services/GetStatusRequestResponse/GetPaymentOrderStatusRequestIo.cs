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
    public partial class GetPaymentOrderStatusRequestIo
    {

        private long singlePaymentIdField;

        private bool singlePaymentIdFieldSpecified;

        private long batchPaymentIdField;

        private bool batchPaymentIdFieldSpecified;

        /// <remarks/>
        public long singlePaymentId
        {
            get
            {
                return this.singlePaymentIdField;
            }
            set
            {
                this.singlePaymentIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool singlePaymentIdSpecified
        {
            get
            {
                return this.singlePaymentIdFieldSpecified;
            }
            set
            {
                this.singlePaymentIdFieldSpecified = value;
            }
        }

        /// <remarks/>
        public long batchPaymentId
        {
            get
            {
                return this.batchPaymentIdField;
            }
            set
            {
                this.batchPaymentIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool batchPaymentIdSpecified
        {
            get
            {
                return this.batchPaymentIdFieldSpecified;
            }
            set
            {
                this.batchPaymentIdFieldSpecified = value;
            }
        }
    }

}
