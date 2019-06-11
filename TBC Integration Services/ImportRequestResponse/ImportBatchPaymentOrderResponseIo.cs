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
    public partial class ImportBatchPaymentOrderResponseIo
    {

        private long mygeminiBatchIdField;

        private bool mygeminiBatchIdFieldSpecified;

        /// <remarks/>
        public long mygeminiBatchId
        {
            get
            {
                return this.mygeminiBatchIdField;
            }
            set
            {
                this.mygeminiBatchIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool mygeminiBatchIdSpecified
        {
            get
            {
                return this.mygeminiBatchIdFieldSpecified;
            }
            set
            {
                this.mygeminiBatchIdFieldSpecified = value;
            }
        }
    }

}
