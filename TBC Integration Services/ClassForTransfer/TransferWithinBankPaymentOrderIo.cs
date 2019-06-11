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
    public partial class TransferWithinBankPaymentOrderIo : PaymentOrderIo
    {

        private string beneficiaryNameField;

        private string beneficiaryTaxCodeField;

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
        public string beneficiaryTaxCode
        {
            get
            {
                return this.beneficiaryTaxCodeField;
            }
            set
            {
                this.beneficiaryTaxCodeField = value;
            }
        }
    }

}
