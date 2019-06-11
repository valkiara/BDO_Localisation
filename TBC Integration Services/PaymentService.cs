using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Net;
using System.Web.Services;
using System.Web.Services.Description;
using System.Web.Services.Protocols;
using System.Xml.Serialization;


namespace BDO_Localisation_AddOn.TBC_Integration_Services
{
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name = "PaymentServiceBinding", Namespace = "http://www.mygemini.com/schemas/mygemini")]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(AbstractIo))]
    public partial class PaymentService : SoapHttpClientProtocol
    {

        private System.Threading.SendOrPostCallback ImportSinglePaymentOrdersOperationCompleted;

        private System.Threading.SendOrPostCallback ImportBatchPaymentOrderOperationCompleted;

        private System.Threading.SendOrPostCallback GetPaymentOrderStatusOperationCompleted;

        /// <remarks/>
        public PaymentService()
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
        public event ImportSinglePaymentOrdersCompletedEventHandler ImportSinglePaymentOrdersCompleted;

        /// <remarks/>
        public event ImportBatchPaymentOrderCompletedEventHandler ImportBatchPaymentOrderCompleted;

        /// <remarks/>
        public event GetPaymentOrderStatusCompletedEventHandler GetPaymentOrderStatusCompleted;

        public Security secHeader { get; set; }

        /// <remarks/>
        [SoapHeader("secHeader", Direction = SoapHeaderDirection.In)]
        [SoapDocumentMethod("http://www.mygemini.com/schemas/mygemini/ImportSinglePaymentOrders", ParameterStyle = SoapParameterStyle.Bare, Use = SoapBindingUse.Literal)]
        [return: XmlElement("ImportSinglePaymentOrdersResponseIo", Namespace = "http://www.mygemini.com/schemas/mygemini")]
        public ImportSinglePaymentOrdersResponseIo ImportSinglePaymentOrders([XmlElement(Namespace = "http://www.mygemini.com/schemas/mygemini")] ImportSinglePaymentOrdersRequestIo ImportSinglePaymentOrdersRequestIo)
        {
            return (ImportSinglePaymentOrdersResponseIo)this.Invoke("ImportSinglePaymentOrders", new object[1]
            {
        (object) ImportSinglePaymentOrdersRequestIo
            })[0];
        }

        /// <remarks/>
        public System.IAsyncResult BeginImportSinglePaymentOrders(PaymentOrderIo[] ImportSinglePaymentOrdersRequestIo, System.AsyncCallback callback, object asyncState)
        {
            return this.BeginInvoke("ImportSinglePaymentOrders", new object[] {
                    ImportSinglePaymentOrdersRequestIo}, callback, asyncState);
        }

        /// <remarks/>
        public PaymentOrderResultIo[] EndImportSinglePaymentOrders(System.IAsyncResult asyncResult)
        {
            object[] results = this.EndInvoke(asyncResult);
            return ((PaymentOrderResultIo[])(results[0]));
        }

        /// <remarks/>
        public void ImportSinglePaymentOrdersAsync(PaymentOrderIo[] ImportSinglePaymentOrdersRequestIo)
        {
            this.ImportSinglePaymentOrdersAsync(ImportSinglePaymentOrdersRequestIo, null);
        }

        /// <remarks/>
        public void ImportSinglePaymentOrdersAsync(PaymentOrderIo[] ImportSinglePaymentOrdersRequestIo, object userState)
        {
            if ((this.ImportSinglePaymentOrdersOperationCompleted == null))
            {
                this.ImportSinglePaymentOrdersOperationCompleted = new System.Threading.SendOrPostCallback(this.OnImportSinglePaymentOrdersOperationCompleted);
            }
            this.InvokeAsync("ImportSinglePaymentOrders", new object[] {
                    ImportSinglePaymentOrdersRequestIo}, this.ImportSinglePaymentOrdersOperationCompleted, userState);
        }

        private void OnImportSinglePaymentOrdersOperationCompleted(object arg)
        {
            if ((this.ImportSinglePaymentOrdersCompleted != null))
            {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.ImportSinglePaymentOrdersCompleted(this, new ImportSinglePaymentOrdersCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }

        /// <remarks/>
        [SoapHeader("secHeader", Direction = SoapHeaderDirection.In)]
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.mygemini.com/schemas/mygemini/ImportBatchPaymentOrder", Use = System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle = System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("ImportBatchPaymentOrderResponseIo", Namespace = "http://www.mygemini.com/schemas/mygemini")]
        public ImportBatchPaymentOrderResponseIo ImportBatchPaymentOrder([System.Xml.Serialization.XmlElementAttribute(Namespace = "http://www.mygemini.com/schemas/mygemini")] ImportBatchPaymentOrderRequestIo ImportBatchPaymentOrderRequestIo)
        {
            object[] results = this.Invoke("ImportBatchPaymentOrder", new object[] {
                    ImportBatchPaymentOrderRequestIo});
            return ((ImportBatchPaymentOrderResponseIo)(results[0]));
        }

        /// <remarks/>
        public System.IAsyncResult BeginImportBatchPaymentOrder(ImportBatchPaymentOrderRequestIo ImportBatchPaymentOrderRequestIo, System.AsyncCallback callback, object asyncState)
        {
            return this.BeginInvoke("ImportBatchPaymentOrder", new object[] {
                    ImportBatchPaymentOrderRequestIo}, callback, asyncState);
        }

        /// <remarks/>
        public ImportBatchPaymentOrderResponseIo EndImportBatchPaymentOrder(System.IAsyncResult asyncResult)
        {
            object[] results = this.EndInvoke(asyncResult);
            return ((ImportBatchPaymentOrderResponseIo)(results[0]));
        }

        /// <remarks/>
        public void ImportBatchPaymentOrderAsync(ImportBatchPaymentOrderRequestIo ImportBatchPaymentOrderRequestIo)
        {
            this.ImportBatchPaymentOrderAsync(ImportBatchPaymentOrderRequestIo, null);
        }

        /// <remarks/>
        public void ImportBatchPaymentOrderAsync(ImportBatchPaymentOrderRequestIo ImportBatchPaymentOrderRequestIo, object userState)
        {
            if ((this.ImportBatchPaymentOrderOperationCompleted == null))
            {
                this.ImportBatchPaymentOrderOperationCompleted = new System.Threading.SendOrPostCallback(this.OnImportBatchPaymentOrderOperationCompleted);
            }
            this.InvokeAsync("ImportBatchPaymentOrder", new object[] {
                    ImportBatchPaymentOrderRequestIo}, this.ImportBatchPaymentOrderOperationCompleted, userState);
        }

        private void OnImportBatchPaymentOrderOperationCompleted(object arg)
        {
            if ((this.ImportBatchPaymentOrderCompleted != null))
            {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.ImportBatchPaymentOrderCompleted(this, new ImportBatchPaymentOrderCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }

        /// <remarks/>
        [SoapHeader("secHeader", Direction = SoapHeaderDirection.In)]
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.mygemini.com/schemas/mygemini/GetPaymentOrderStatus", Use = System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle = System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("GetPaymentOrderStatusResponseIo", Namespace = "http://www.mygemini.com/schemas/mygemini")]
        public GetPaymentOrderStatusResponseIo GetPaymentOrderStatus([System.Xml.Serialization.XmlElementAttribute(Namespace = "http://www.mygemini.com/schemas/mygemini")] GetPaymentOrderStatusRequestIo GetPaymentOrderStatusRequestIo)
        {
            object[] results = this.Invoke("GetPaymentOrderStatus", new object[] {
                    GetPaymentOrderStatusRequestIo});
            return ((GetPaymentOrderStatusResponseIo)(results[0]));
        }

        /// <remarks/>
        public System.IAsyncResult BeginGetPaymentOrderStatus(GetPaymentOrderStatusRequestIo GetPaymentOrderStatusRequestIo, System.AsyncCallback callback, object asyncState)
        {
            return this.BeginInvoke("GetPaymentOrderStatus", new object[] {
                    GetPaymentOrderStatusRequestIo}, callback, asyncState);
        }

        /// <remarks/>
        public GetPaymentOrderStatusResponseIo EndGetPaymentOrderStatus(System.IAsyncResult asyncResult)
        {
            object[] results = this.EndInvoke(asyncResult);
            return ((GetPaymentOrderStatusResponseIo)(results[0]));
        }

        /// <remarks/>
        public void GetPaymentOrderStatusAsync(GetPaymentOrderStatusRequestIo GetPaymentOrderStatusRequestIo)
        {
            this.GetPaymentOrderStatusAsync(GetPaymentOrderStatusRequestIo, null);
        }

        /// <remarks/>
        public void GetPaymentOrderStatusAsync(GetPaymentOrderStatusRequestIo GetPaymentOrderStatusRequestIo, object userState)
        {
            if ((this.GetPaymentOrderStatusOperationCompleted == null))
            {
                this.GetPaymentOrderStatusOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetPaymentOrderStatusOperationCompleted);
            }
            this.InvokeAsync("GetPaymentOrderStatus", new object[] {
                    GetPaymentOrderStatusRequestIo}, this.GetPaymentOrderStatusOperationCompleted, userState);
        }

        private void OnGetPaymentOrderStatusOperationCompleted(object arg)
        {
            if ((this.GetPaymentOrderStatusCompleted != null))
            {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetPaymentOrderStatusCompleted(this, new GetPaymentOrderStatusCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
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
    public delegate void ImportSinglePaymentOrdersCompletedEventHandler(object sender, ImportSinglePaymentOrdersCompletedEventArgs e);

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class ImportSinglePaymentOrdersCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs
    {

        private object[] results;

        internal ImportSinglePaymentOrdersCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) :
            base(exception, cancelled, userState)
        {
            this.results = results;
        }

        /// <remarks/>
        public PaymentOrderResultIo[] Result
        {
            get
            {
                this.RaiseExceptionIfNecessary();
                return ((PaymentOrderResultIo[])(this.results[0]));
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    public delegate void ImportBatchPaymentOrderCompletedEventHandler(object sender, ImportBatchPaymentOrderCompletedEventArgs e);

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class ImportBatchPaymentOrderCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs
    {

        private object[] results;

        internal ImportBatchPaymentOrderCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) :
            base(exception, cancelled, userState)
        {
            this.results = results;
        }

        /// <remarks/>
        public ImportBatchPaymentOrderResponseIo Result
        {
            get
            {
                this.RaiseExceptionIfNecessary();
                return ((ImportBatchPaymentOrderResponseIo)(this.results[0]));
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    public delegate void GetPaymentOrderStatusCompletedEventHandler(object sender, GetPaymentOrderStatusCompletedEventArgs e);

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.0.30319.33440")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetPaymentOrderStatusCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs
    {

        private object[] results;

        internal GetPaymentOrderStatusCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) :
            base(exception, cancelled, userState)
        {
            this.results = results;
        }

        /// <remarks/>
        public GetPaymentOrderStatusResponseIo Result
        {
            get
            {
                this.RaiseExceptionIfNecessary();
                return ((GetPaymentOrderStatusResponseIo)(this.results[0]));
            }
        }
    }
}
