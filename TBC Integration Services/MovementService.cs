using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Net;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Xml.Serialization;

namespace BDO_Localisation_AddOn.TBC_Integration_Services
{
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name = "MovementServiceBinding", Namespace = "http://www.mygemini.com/schemas/mygemini")]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(AbstractIo))]
    public partial class MovementService : System.Web.Services.Protocols.SoapHttpClientProtocol
    {

        private System.Threading.SendOrPostCallback GetAccountMovementsOperationCompleted;

        /// <remarks/>
        public MovementService()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            this.Url = "http://localhost:8080/dbi/dbiService";
        }

        public void SetUrl(string ServiceUrl)
        {
            this.Url = ServiceUrl;
        }

        public void SetUsernameToken(string Username, string Password, string Nonce)
        {
            this.secHeader = new Security();
            this.secHeader.UsernameToken = new UsernameToken();
            this.secHeader.UsernameToken.Username = Username;
            this.secHeader.UsernameToken.Password = Password;
            this.secHeader.UsernameToken.Nonce = Nonce;
        }

        /// <remarks/>
        public event GetAccountMovementsCompletedEventHandler GetAccountMovementsCompleted;

        public Security secHeader { get; set; }

        /// <remarks/>
        [SoapHeader("secHeader", Direction = SoapHeaderDirection.In)]
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.mygemini.com/schemas/mygemini/GetAccountMovements", RequestElementName = "GetAccountMovementsRequestIo", RequestNamespace = "http://www.mygemini.com/schemas/mygemini", ResponseElementName = "GetAccountMovementsResponseIo", ResponseNamespace = "http://www.mygemini.com/schemas/mygemini", Use = System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle = System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        [return: System.Xml.Serialization.XmlElementAttribute("result")]
        public BaseQueryResultIo GetAccountMovements(AccountMovementFilterIo accountMovementFilterIo, [System.Xml.Serialization.XmlElementAttribute("accountMovement")] out AccountMovementDetailIo[] accountMovement)
        {
            object[] results = this.Invoke("GetAccountMovements", new object[] {
                    accountMovementFilterIo});
            accountMovement = ((AccountMovementDetailIo[])(results[1]));
            return ((BaseQueryResultIo)(results[0]));
        }

        /// <remarks/>
        public System.IAsyncResult BeginGetAccountMovements(AccountMovementFilterIo accountMovementFilterIo, System.AsyncCallback callback, object asyncState)
        {
            return this.BeginInvoke("GetAccountMovements", new object[] {
                    accountMovementFilterIo}, callback, asyncState);
        }

        /// <remarks/>
        public BaseQueryResultIo EndGetAccountMovements(System.IAsyncResult asyncResult, out AccountMovementDetailIo[] accountMovement)
        {
            object[] results = this.EndInvoke(asyncResult);
            accountMovement = ((AccountMovementDetailIo[])(results[1]));
            return ((BaseQueryResultIo)(results[0]));
        }

        /// <remarks/>
        public void GetAccountMovementsAsync(AccountMovementFilterIo accountMovementFilterIo)
        {
            this.GetAccountMovementsAsync(accountMovementFilterIo, null);
        }

        /// <remarks/>
        public void GetAccountMovementsAsync(AccountMovementFilterIo accountMovementFilterIo, object userState)
        {
            if ((this.GetAccountMovementsOperationCompleted == null))
            {
                this.GetAccountMovementsOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetAccountMovementsOperationCompleted);
            }
            this.InvokeAsync("GetAccountMovements", new object[] {
                    accountMovementFilterIo}, this.GetAccountMovementsOperationCompleted, userState);
        }

        private void OnGetAccountMovementsOperationCompleted(object arg)
        {
            if ((this.GetAccountMovementsCompleted != null))
            {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetAccountMovementsCompleted(this, new GetAccountMovementsCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }

        /// <remarks/>
        public new void CancelAsync(object userState)
        {
            base.CancelAsync(userState);
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    public delegate void GetAccountMovementsCompletedEventHandler(object sender, GetAccountMovementsCompletedEventArgs e);

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetAccountMovementsCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs
    {

        private object[] results;

        internal GetAccountMovementsCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) :
            base(exception, cancelled, userState)
        {
            this.results = results;
        }

        /// <remarks/>
        public BaseQueryResultIo Result
        {
            get
            {
                this.RaiseExceptionIfNecessary();
                return ((BaseQueryResultIo)(this.results[0]));
            }
        }

        /// <remarks/>
        public AccountMovementDetailIo[] accountMovement
        {
            get
            {
                this.RaiseExceptionIfNecessary();
                return ((AccountMovementDetailIo[])(this.results[1]));
            }
        }
    }
}
