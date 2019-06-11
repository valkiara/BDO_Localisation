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
    [System.Web.Services.WebServiceBindingAttribute(Name = "ChangePasswordServiceBinding", Namespace = "http://www.mygemini.com/schemas/mygemini")]
    public partial class ChangePasswordService : System.Web.Services.Protocols.SoapHttpClientProtocol
    {

        private System.Threading.SendOrPostCallback ChangePasswordOperationCompleted;

        /// <remarks/>
        public ChangePasswordService()
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
        public event ChangePasswordCompletedEventHandler ChangePasswordCompleted;

        public Security secHeader { get; set; }

        /// <remarks/>
        [SoapHeader("secHeader", Direction = SoapHeaderDirection.In)]
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.mygemini.com/schemas/mygemini/ChangePassword", RequestElementName = "ChangePasswordRequestIo", RequestNamespace = "http://www.mygemini.com/schemas/mygemini", ResponseElementName = "ChangePasswordResponseIo", ResponseNamespace = "http://www.mygemini.com/schemas/mygemini", Use = System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle = System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        [return: System.Xml.Serialization.XmlElementAttribute("message")]
        public string ChangePassword(string newPassword)
        {
            object[] results = this.Invoke("ChangePassword", new object[] {
                    newPassword});
            return ((string)(results[0]));
        }

        /// <remarks/>
        public System.IAsyncResult BeginChangePassword(string newPassword, System.AsyncCallback callback, object asyncState)
        {
            return this.BeginInvoke("ChangePassword", new object[] {
                    newPassword}, callback, asyncState);
        }

        /// <remarks/>
        public string EndChangePassword(System.IAsyncResult asyncResult)
        {
            object[] results = this.EndInvoke(asyncResult);
            return ((string)(results[0]));
        }

        /// <remarks/>
        public void ChangePasswordAsync(string newPassword)
        {
            this.ChangePasswordAsync(newPassword, null);
        }

        /// <remarks/>
        public void ChangePasswordAsync(string newPassword, object userState)
        {
            if ((this.ChangePasswordOperationCompleted == null))
            {
                this.ChangePasswordOperationCompleted = new System.Threading.SendOrPostCallback(this.OnChangePasswordOperationCompleted);
            }
            this.InvokeAsync("ChangePassword", new object[] {
                    newPassword}, this.ChangePasswordOperationCompleted, userState);
        }

        private void OnChangePasswordOperationCompleted(object arg)
        {
            if ((this.ChangePasswordCompleted != null))
            {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.ChangePasswordCompleted(this, new ChangePasswordCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
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
    public delegate void ChangePasswordCompletedEventHandler(object sender, ChangePasswordCompletedEventArgs e);

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class ChangePasswordCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs
    {

        private object[] results;

        internal ChangePasswordCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) :
            base(exception, cancelled, userState)
        {
            this.results = results;
        }

        /// <remarks/>
        public string Result
        {
            get
            {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
}