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
    public partial class AccountMovementFilterIo : BaseFilterIo
    {

        private string accountNumberField;

        private string accountCurrencyCodeField;

        private System.DateTime periodFromField;

        private bool periodFromFieldSpecified;

        private System.DateTime periodToField;

        private bool periodToFieldSpecified;

        private string movementIdField;

        private System.DateTime lastMovementTimestampField;

        private bool lastMovementTimestampFieldSpecified;

        /// <remarks/>
        public string accountNumber
        {
            get
            {
                return this.accountNumberField;
            }
            set
            {
                this.accountNumberField = value;
            }
        }

        /// <remarks/>
        public string accountCurrencyCode
        {
            get
            {
                return this.accountCurrencyCodeField;
            }
            set
            {
                this.accountCurrencyCodeField = value;
            }
        }

        /// <remarks/>
        public System.DateTime periodFrom
        {
            get
            {
                return this.periodFromField;
            }
            set
            {
                this.periodFromField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool periodFromSpecified
        {
            get
            {
                return this.periodFromFieldSpecified;
            }
            set
            {
                this.periodFromFieldSpecified = value;
            }
        }

        /// <remarks/>
        public System.DateTime periodTo
        {
            get
            {
                return this.periodToField;
            }
            set
            {
                this.periodToField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool periodToSpecified
        {
            get
            {
                return this.periodToFieldSpecified;
            }
            set
            {
                this.periodToFieldSpecified = value;
            }
        }

        /// <remarks/>
        public string movementId
        {
            get
            {
                return this.movementIdField;
            }
            set
            {
                this.movementIdField = value;
            }
        }

        /// <remarks/>
        public System.DateTime lastMovementTimestamp
        {
            get
            {
                return this.lastMovementTimestampField;
            }
            set
            {
                this.lastMovementTimestampField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool lastMovementTimestampSpecified
        {
            get
            {
                return this.lastMovementTimestampFieldSpecified;
            }
            set
            {
                this.lastMovementTimestampFieldSpecified = value;
            }
        }
    }

}
