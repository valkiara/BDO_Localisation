
using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Xml.Serialization;

namespace BDO_Localisation_AddOn.TBC_Integration_Services
{
    [GeneratedCode("wsdl", "4.0.30319.33440")]
    [XmlType(AnonymousType = true, Namespace = "http://www.mygemini.com/schemas/mygemini")]
    [DesignerCategory("code")]
    [DebuggerStepThrough]
    [Serializable]
    public class ImportSinglePaymentOrdersResponseIo : AbstractIo
    {
        private PaymentOrderResultIo[] paymentOrdersResultsField;

        [XmlElement("PaymentOrdersResults")]
        public PaymentOrderResultIo[] PaymentOrdersResults
        {
            get
            {
                return this.paymentOrdersResultsField;
            }
            set
            {
                this.paymentOrdersResultsField = value;
            }
        }

        public new string ToString()
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append("PaymentOrderResults[");
            if (this.paymentOrdersResultsField != null)
            {
                PaymentOrderResultIo[] paymentOrderResultIoArray = this.paymentOrdersResultsField;
                int index = 0;
                while (index < paymentOrderResultIoArray.Length)
                {
                    PaymentOrderResultIo paymentOrderResultIo = paymentOrderResultIoArray[index];
                    stringBuilder.Append(paymentOrderResultIo.ToString()).Append(" ");
                    checked { ++index; }
                }
            }
            else
                stringBuilder.Append("<empty>");
            stringBuilder.Append("]");
            return stringBuilder.ToString();
        }
    }
}
