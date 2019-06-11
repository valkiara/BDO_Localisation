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
    public partial class PaymentStatusDataIo : AbstractIo
    {

        private int positionField;

        private bool positionFieldSpecified;

        private string paymentIdField;

        private string paymentStatusField;

        private string errorDetailENField;

        private string errorDetailGEField;

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
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool positionSpecified
        {
            get
            {
                return this.positionFieldSpecified;
            }
            set
            {
                this.positionFieldSpecified = value;
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
        public string paymentStatus
        {
            get
            {
                return this.paymentStatusField;
            }
            set
            {
                this.paymentStatusField = value;
            }
        }

        /// <remarks/>
        public string errorDetailEN
        {
            get
            {
                return this.errorDetailENField;
            }
            set
            {
                this.errorDetailENField = value;
            }
        }

        /// <remarks/>
        public string errorDetailGE
        {
            get
            {
                return this.errorDetailGEField;
            }
            set
            {
                this.errorDetailGEField = value;
            }
        }
    }
}
