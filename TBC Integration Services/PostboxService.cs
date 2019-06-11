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
    [System.Web.Services.WebServiceBindingAttribute(Name = "PostboxServiceBinding", Namespace = "http://www.mygemini.com/schemas/mygemini")]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(AbstractIo))]
    public partial class PostboxService : System.Web.Services.Protocols.SoapHttpClientProtocol
    {

        private System.Threading.SendOrPostCallback GetMessagesFromPostboxOperationCompleted;

        private System.Threading.SendOrPostCallback AcknowledgePostboxMessagesOperationCompleted;

        /// <remarks/>
        public PostboxService()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            this.Url = "http://localhost:8080/dbi/dbiService";
        }

        public void SetUrl(string ServiceUrl)
        {
            this.Url = ServiceUrl;
        }

        public void SetUsernameToken(string Username, string Password)
        {
            this.secHeader = new Security();
            this.secHeader.UsernameToken = new UsernameToken();
            this.secHeader.UsernameToken.Username = Username;
            this.secHeader.UsernameToken.Password = Password;
        }

        public Security secHeader { get; set; }

        /// <remarks/>
        public event GetMessagesFromPostboxCompletedEventHandler GetMessagesFromPostboxCompleted;

        /// <remarks/>
        public event AcknowledgePostboxMessagesCompletedEventHandler AcknowledgePostboxMessagesCompleted;

        /// <remarks/>
        [SoapHeader("secHeader", Direction = SoapHeaderDirection.In)]
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.mygemini.com/schemas/mygemini/GetMessagesFromPostbox", RequestElementName = "GetPostboxMessagesRequestIo", RequestNamespace = "http://www.mygemini.com/schemas/mygemini", ResponseElementName = "GetPostboxMessagesResponseIo", ResponseNamespace = "http://www.mygemini.com/schemas/mygemini", Use = System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle = System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        [return: System.Xml.Serialization.XmlElementAttribute("messages")]
        public PostboxMessageIo[] GetMessagesFromPostbox(string messageType)
        {
            object[] results = this.Invoke("GetMessagesFromPostbox", new object[] {
                    messageType});
            return ((PostboxMessageIo[])(results[0]));
        }

        /// <remarks/>
        public System.IAsyncResult BeginGetMessagesFromPostbox(string messageType, System.AsyncCallback callback, object asyncState)
        {
            return this.BeginInvoke("GetMessagesFromPostbox", new object[] {
                    messageType}, callback, asyncState);
        }

        /// <remarks/>
        public PostboxMessageIo[] EndGetMessagesFromPostbox(System.IAsyncResult asyncResult)
        {
            object[] results = this.EndInvoke(asyncResult);
            return ((PostboxMessageIo[])(results[0]));
        }

        /// <remarks/>
        public void GetMessagesFromPostboxAsync(string messageType)
        {
            this.GetMessagesFromPostboxAsync(messageType, null);
        }

        /// <remarks/>
        public void GetMessagesFromPostboxAsync(string messageType, object userState)
        {
            if ((this.GetMessagesFromPostboxOperationCompleted == null))
            {
                this.GetMessagesFromPostboxOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetMessagesFromPostboxOperationCompleted);
            }
            this.InvokeAsync("GetMessagesFromPostbox", new object[] {
                    messageType}, this.GetMessagesFromPostboxOperationCompleted, userState);
        }

        private void OnGetMessagesFromPostboxOperationCompleted(object arg)
        {
            if ((this.GetMessagesFromPostboxCompleted != null))
            {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetMessagesFromPostboxCompleted(this, new GetMessagesFromPostboxCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }

        /// <remarks/>
        [SoapHeader("secHeader", Direction = SoapHeaderDirection.In)]
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.mygemini.com/schemas/mygemini/AcknowledgePostboxMessages", RequestElementName = "PostboxAcknowledgementRequestIo", RequestNamespace = "http://www.mygemini.com/schemas/mygemini", ResponseElementName = "PostboxAcknowledgementResponseIo", ResponseNamespace = "http://www.mygemini.com/schemas/mygemini", Use = System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle = System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        [return: System.Xml.Serialization.XmlElementAttribute("responseText")]
        public string AcknowledgePostboxMessages([System.Xml.Serialization.XmlElementAttribute("messageIds")] long[] messageIds)
        {
            object[] results = this.Invoke("AcknowledgePostboxMessages", new object[] {
                    messageIds});
            return ((string)(results[0]));
        }

        /// <remarks/>
        public System.IAsyncResult BeginAcknowledgePostboxMessages(long[] messageIds, System.AsyncCallback callback, object asyncState)
        {
            return this.BeginInvoke("AcknowledgePostboxMessages", new object[] {
                    messageIds}, callback, asyncState);
        }

        /// <remarks/>
        public string EndAcknowledgePostboxMessages(System.IAsyncResult asyncResult)
        {
            object[] results = this.EndInvoke(asyncResult);
            return ((string)(results[0]));
        }

        /// <remarks/>
        public void AcknowledgePostboxMessagesAsync(long[] messageIds)
        {
            this.AcknowledgePostboxMessagesAsync(messageIds, null);
        }

        /// <remarks/>
        public void AcknowledgePostboxMessagesAsync(long[] messageIds, object userState)
        {
            if ((this.AcknowledgePostboxMessagesOperationCompleted == null))
            {
                this.AcknowledgePostboxMessagesOperationCompleted = new System.Threading.SendOrPostCallback(this.OnAcknowledgePostboxMessagesOperationCompleted);
            }
            this.InvokeAsync("AcknowledgePostboxMessages", new object[] {
                    messageIds}, this.AcknowledgePostboxMessagesOperationCompleted, userState);
        }

        private void OnAcknowledgePostboxMessagesOperationCompleted(object arg)
        {
            if ((this.AcknowledgePostboxMessagesCompleted != null))
            {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.AcknowledgePostboxMessagesCompleted(this, new AcknowledgePostboxMessagesCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
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
    public delegate void GetMessagesFromPostboxCompletedEventHandler(object sender, GetMessagesFromPostboxCompletedEventArgs e);

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetMessagesFromPostboxCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs
    {

        private object[] results;

        internal GetMessagesFromPostboxCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) :
            base(exception, cancelled, userState)
        {
            this.results = results;
        }

        /// <remarks/>
        public PostboxMessageIo[] Result
        {
            get
            {
                this.RaiseExceptionIfNecessary();
                return ((PostboxMessageIo[])(this.results[0]));
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    public delegate void AcknowledgePostboxMessagesCompletedEventHandler(object sender, AcknowledgePostboxMessagesCompletedEventArgs e);

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class AcknowledgePostboxMessagesCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs
    {

        private object[] results;

        internal AcknowledgePostboxMessagesCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) :
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
