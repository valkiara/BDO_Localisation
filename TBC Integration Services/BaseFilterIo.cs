using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn.TBC_Integration_Services
{
    /// <remarks/>
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(AccountMovementFilterIo))]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://www.mygemini.com/schemas/mygemini")]
    public partial class BaseFilterIo : AbstractIo
    {

        private BasePagerIo pagerField;

        private AdditionalAttributeIo[] additionalAttributesField;

        /// <remarks/>
        public BasePagerIo pager
        {
            get
            {
                return this.pagerField;
            }
            set
            {
                this.pagerField = value;
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
