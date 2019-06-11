
using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Xml.Serialization;

namespace BDO_Localisation_AddOn.TBC_Integration_Services
{
    [DebuggerStepThrough]
    [XmlType(AnonymousType = true, Namespace = "http://www.mygemini.com/schemas/mygemini")]
    [DesignerCategory("code")]
    [GeneratedCode("wsdl", "4.0.30319.33440")]
    [Serializable]
    public class ImportSinglePaymentOrdersRequestIo : AbstractIo
    {
        private PaymentOrderIo[] singlePaymentOrderField;

        [XmlElement("singlePaymentOrder")]
        public PaymentOrderIo[] singlePaymentOrder
        {
            get
            {
                return this.singlePaymentOrderField;
            }
            set
            {
                this.singlePaymentOrderField = value;
            }
        }
    }
}
