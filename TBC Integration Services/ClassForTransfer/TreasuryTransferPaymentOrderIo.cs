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
    public partial class TreasuryTransferPaymentOrderIo : PaymentOrderIo
    {

        private string taxpayerCodeField;

        private string taxpayerNameField;

        private string treasuryCodeField;

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
    }

}
