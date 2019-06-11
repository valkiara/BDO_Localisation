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
    public partial class PostboxMessageIo : AbstractIo
    {

        private long messageIdField;

        private string messageTextField;

        private string messageTypeField;

        private string messageStatusField;

        private AdditionalAttributeIo[] additionalAttributesField;

        /// <remarks/>
        public long messageId
        {
            get
            {
                return this.messageIdField;
            }
            set
            {
                this.messageIdField = value;
            }
        }

        /// <remarks/>
        public string messageText
        {
            get
            {
                return this.messageTextField;
            }
            set
            {
                this.messageTextField = value;
            }
        }

        /// <remarks/>
        public string messageType
        {
            get
            {
                return this.messageTypeField;
            }
            set
            {
                this.messageTypeField = value;
            }
        }

        /// <remarks/>
        public string messageStatus
        {
            get
            {
                return this.messageStatusField;
            }
            set
            {
                this.messageStatusField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("additionalAttributes")]
        public AdditionalAttributeIo[] additionalAttributes
        {
            get
            {
                return this.additionalAttributesField;
            }
            set
            {
                this.additionalAttributesField = value;
            }
        }
    }
}
