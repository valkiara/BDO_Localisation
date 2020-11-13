using BDO_Localisation_AddOn.BOG_Integration_Services.Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BDO_Localisation_AddOn.BOG_Integration_Services
{
    static partial class BDOSAuthenticationFormBOG
    {
        private static SAPbouiCOM.Form form { get; set; }
        private static string operation { get; set; }
        private static List<int> docEntryList { get; set; }
        private static bool importBatchPaymentOrders { get; set; }
        private static string batchName { get; set; }
        //---> BOG
        private static string AuthorizeUrl { get; set; }
        private static string CallbackUri { get; set; }
        private static string ClientId { get; set; }
        private static AuthorizeResponse AuthorizeResponse { get; set; }
        private static string ApiBaseUrl { get; set; }
        private static bool LocationURL { get; set; }
        private static StatementFilter oStatementFilter { get; set; }
        //<--- BOG

        public static void createForm(SAPbouiCOM.Form formOutgoingPayment, string operationOutgoingPayment, List<int> docEntryListOutgoingPayment, bool importBatchPaymentOrdersOutgoingPayment, string batchNameOutgoingPayment, StatementFilter oStatementFilterTmp, out string errorText)
        {
            errorText = null;
            form = formOutgoingPayment;
            operation = operationOutgoingPayment;
            docEntryList = docEntryListOutgoingPayment;
            importBatchPaymentOrders = importBatchPaymentOrdersOutgoingPayment;
            batchName = batchNameOutgoingPayment;
            oStatementFilter = oStatementFilterTmp;

            //---> BOG
            string clientIdTemp;
            int port;
            string mode;
            string url;
            string wsdl = CommonFunctions.getServiceWSDLForInternetBanking("BOG", out clientIdTemp, out port, out mode, out url, out errorText);
            string scope = "read write";

            if (string.IsNullOrEmpty(errorText) == false)
            {
                return;
            }

            if (mode == "test") //სატესტო გარემოში port შევსებული უნდა იყოს
            {
                AuthorizeUrl = $"{url}:{port}/Oauth/Connect/Authorize.aspx";
                CallbackUri = $"{url}:{port}/Oauth/Connect/Token.aspx";

            }
            else if (mode == "real")
            {
                AuthorizeUrl = $"{url}/Oauth/Connect/Authorize.aspx";
                CallbackUri = $"{url}/Oauth/Connect/Token.aspx";
            }
            else if (mode == "realNew")
            {
                AuthorizeUrl = $"{url}/auth/realms/bog/protocol/openid-connect/auth";
                CallbackUri = $"{url}/auth/realms/bog/protocol/openid-connect/token";
                scope = "corp";
            }
            else if (mode == "testNew")
            {
                AuthorizeUrl = $"{url}/auth/realms/bog-test/protocol/openid-connect/auth";
                CallbackUri = $"{url}/auth/realms/bog-test/protocol/openid-connect/token";
                scope = "corp";
            }

            ClientId = clientIdTemp;
            ApiBaseUrl = wsdl;
            LocationURL = false;

            var client = new OAuth2Client(new Uri(AuthorizeUrl));
            var state = Guid.NewGuid().ToString();
            var startUrl = client.CreateAuthorizeUrl(ClientId, "token", scope, CallbackUri, state);
            //---> BOG

            int formHeight = Program.uiApp.Desktop.Width;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSAuthenticationFormBOG");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("LoginRequired"));
            //formProperties.Add("Left", (Program.uiApp.Desktop.Width - formWidth) / 2);
            formProperties.Add("ClientWidth", formWidth);
            //formProperties.Add("Top", (Program.uiApp.Desktop.Height - formHeight) / 3);
            formProperties.Add("ClientHeight", formHeight);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }

            if (formExist == true)
            {
                if (newForm == true)
                {
                    errorText = null;

                    //SAPbouiCOM.Item oBrowser = oForm.Items.Add("urlWB", SAPbouiCOM.BoFormItemTypes.it_WEB_BROWSER);
                    //oBrowser.Top = 15;
                    //oBrowser.Left = 15;
                    //oBrowser.Width = oForm.Width - 60;
                    //oBrowser.Height = oForm.Height - 120;
                    //SAPbouiCOM.WebBrowser oActive = (SAPbouiCOM.WebBrowser)oBrowser.Specific;
                    //oActive.Url = startUrl;
                    SAPbouiCOM.Item oBrowser = oForm.Items.Add("urlWB", SAPbouiCOM.BoFormItemTypes.it_ACTIVE_X);
                    oBrowser.Top = 15;
                    oBrowser.Left = 15;
                    oBrowser.Width = oForm.Width - 60;
                    oBrowser.Height = oForm.Height - 120;
                    SAPbouiCOM.ActiveX oActive = (SAPbouiCOM.ActiveX)oBrowser.Specific;
                    oActive.ClassID = "Shell.Explorer.2";

                    SHDocVw.WebBrowser WebBrowserChen;
                    WebBrowserChen = (SHDocVw.WebBrowser)oActive.Object;
                    WebBrowserChen.Silent = true;
                    //if (!WebBrowserChen.Busy)
                    //{
                        WebBrowserChen.Navigate2(startUrl);
                        WebBrowserChen.NavigateError += WebBrowserChen_NavigateError;
                        WebBrowserChen.NavigateComplete2 += new SHDocVw.DWebBrowserEvents2_NavigateComplete2EventHandler(myNavigateComplete2);
                        WebBrowserChen.WindowClosing += new SHDocVw.DWebBrowserEvents2_WindowClosingEventHandler(myWindowClosing);
                        WebBrowserChen.WebWorkerFinsihed += new SHDocVw.DWebBrowserEvents2_WebWorkerFinsihedEventHandler(myWebWorkerFinsihed);
                        WebBrowserChen.DocumentComplete += new SHDocVw.DWebBrowserEvents2_DocumentCompleteEventHandler(myDocumentComplete);
                        WebBrowserChen.OnQuit += new SHDocVw.DWebBrowserEvents2_OnQuitEventHandler(myOnQuit);
                        WebBrowserChen.NewProcess += new SHDocVw.DWebBrowserEvents2_NewProcessEventHandler(myNewProcess);
                    //}
                    Dictionary<string, object> formItems;
                    string itemName = "";

                    int left_s = 6;
                    int height = 15;
                    int top = 6;
                    int width_s = 121;

                    top = oForm.ClientHeight - 25;
                    height = height + 4;
                    width_s = 65;

                    itemName = "1";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left_s = left_s + width_s + 2;

                    itemName = "2";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                }
                oForm.Visible = true;
                oForm.Select();
            }
            GC.Collect();
        }

        private static void WebBrowserChen_NavigateError(object pDisp, ref object URL, ref object Frame, ref object StatusCode, ref bool Cancel)
        {
            string k;
            k = "12343124";
            string f = k + "235";
        }

        private static void myNewProcess(int lCauseFlag, object pWB2, ref bool Cancel)
        {
            string k;
            k = "12343124";
            string f = k + "235";
        }

        private static void myOnQuit()
        {

            string k;
            k = "12343124";
            string f = k + "235";
        }

        private static void myDocumentComplete(object pDisp, ref object URL)
        {

            string k;
            k = "12343124";
            string f = k + "235";
        }

        private static void myNavigateComplete2(object pDisp, ref object URL)
        {

            string k;
            k = "12343124";
            string f = k + "235";
        }
        private static void myWindowClosing(bool IsChildWindow, ref bool Cancel)
        {

            string k;
            k = "12343124";
            string f = k + "235";

        }

        private static void myWebWorkerFinsihed(uint dwUniqueID)
        {

            string k;
            k = "12343124";
            string f = k + "235";

        }

        public static void formClose(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (operation == "import") //იმპორტი
                {
                    AssertToken(out errorText);

                    if (string.IsNullOrEmpty(errorText) == false)
                    {
                        Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                    else
                    {
                        HttpClient client = InitializeClient();

                        List<string> infoList = OutgoingPayment.importPaymentOrderBOG(client, docEntryList, importBatchPaymentOrders, batchName, out errorText);
                        if (errorText != null)
                        {
                            Program.uiApp.MessageBox(errorText);
                        }
                        else
                        {
                            for (int i = 0; i < infoList.Count; i++)
                            {
                                Program.uiApp.SetStatusBarMessage(infoList[i], SAPbouiCOM.BoMessageTime.bmt_Short, false);
                            }

                            if (form != null && form.UniqueID == "BDOSInternetBankingForm") //თუ დამუშავებიდან გამოიძახება
                            {
                                BDOSInternetBanking.fillImportMTR(form, out errorText);
                            }
                        }
                    }
                }

                else if (operation == "updateStatus") //სტატუსის განახლება
                {
                    AssertToken(out errorText);

                    if (string.IsNullOrEmpty(errorText) == false)
                    {
                        Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                    else
                    {
                        HttpClient client = InitializeClient();

                        List<string> infoList = OutgoingPayment.updateStatusPaymentOrderBOG(client, docEntryList, out errorText);
                        if (errorText != null)
                        {
                            Program.uiApp.MessageBox(errorText);
                        }
                        else
                        {
                            for (int i = 0; i < infoList.Count; i++)
                            {
                                Program.uiApp.SetStatusBarMessage(infoList[i], SAPbouiCOM.BoMessageTime.bmt_Short, false);
                            }
                            if (form != null && form.UniqueID == "BDOSInternetBankingForm") //თუ დამუშავებიდან გამოიძახება
                            {
                                BDOSInternetBanking.fillImportMTR(form, out errorText);
                            }
                        }
                    }
                }

                else if (operation == "getData") //ამონაწერი
                {
                    HttpClient client = InitializeClient();
                    Task<Statement> oStatement = null;
                    
                    List<StatementDetail> oStatementDetail = null;
                    Task<List<StatementDetail>> oStatementDetailIdTask = null;
                    List<StatementDetail> oStatementDetailId = null;

                    if (oStatementFilter.Page == 0)
                    {
                        oStatement = MainPaymentServiceBOG.getStatement(client, oStatementFilter.AccountNumber, oStatementFilter.Currency, oStatementFilter.PeriodFrom, oStatementFilter.PeriodTo);
                        if (oStatement != null)
                        {
                            oStatementDetail = oStatement.Result.Records;

                            int id = Convert.ToInt32(oStatement.Result.Id);
                            int count = oStatement.Result.Count;

                           

                            if (count > 1000)
                            {

                                double count10000 = count / 1000;
                                double Page = 2;
                                while (Page <= count10000+1)
                                {

                                    oStatementDetailIdTask = MainPaymentServiceBOG.getStatementByID(client, oStatementFilter.AccountNumber, oStatementFilter.Currency, id, Convert.ToInt32(Page));
                                    if (oStatementDetailIdTask != null)
                                    {
                                        oStatementDetailId = oStatementDetailIdTask.Result;
                    }

                                    for (int rowIndex = 0; rowIndex < oStatementDetailId.Count; rowIndex++)
                                    {
                                        StatementDetail newRow = new StatementDetail();

                                        newRow.BeneficiaryDetails = oStatementDetailId[rowIndex].BeneficiaryDetails;
                                        newRow.DocumentActualDate = oStatementDetailId[rowIndex].DocumentActualDate;
                                        newRow.DocumentBeneficiaryInstitution = oStatementDetailId[rowIndex].DocumentBeneficiaryInstitution;
                                        newRow.DocumentBranch = oStatementDetailId[rowIndex].DocumentBranch;
                                        newRow.DocumentCorrespondentAccountNumber = oStatementDetailId[rowIndex].DocumentCorrespondentAccountNumber;
                                        newRow.DocumentCorrespondentBankCode = oStatementDetailId[rowIndex].DocumentCorrespondentBankCode;
                                        newRow.DocumentCorrespondentBankName = oStatementDetailId[rowIndex].DocumentCorrespondentBankName;
                                        newRow.DocumentDestinationAmount = oStatementDetailId[rowIndex].DocumentDestinationAmount;
                                        newRow.DocumentDestinationCurrency = oStatementDetailId[rowIndex].DocumentDestinationCurrency;
                                        newRow.DocumentExpiryDate = oStatementDetailId[rowIndex].DocumentExpiryDate;
                                        newRow.DocumentInformation = oStatementDetailId[rowIndex].DocumentInformation;
                                        newRow.DocumentIntermediaryInstitution = oStatementDetailId[rowIndex].DocumentIntermediaryInstitution;
                                        newRow.DocumentNomination = oStatementDetailId[rowIndex].DocumentNomination;
                                        newRow.DocumentPayee = oStatementDetailId[rowIndex].DocumentPayee;
                                        newRow.DocumentProductGroup = oStatementDetailId[rowIndex].DocumentProductGroup;
                                        newRow.DocumentRate = oStatementDetailId[rowIndex].DocumentRate;
                                        newRow.DocumentRateLimit = oStatementDetailId[rowIndex].DocumentRateLimit;
                                        newRow.DocumentReceiveDate = oStatementDetailId[rowIndex].DocumentReceiveDate;
                                        newRow.DocumentRegistrationRate = oStatementDetailId[rowIndex].DocumentRegistrationRate;
                                        newRow.DocumentSenderInstitution = oStatementDetailId[rowIndex].DocumentSenderInstitution;
                                        newRow.DocumentSourceAmount = oStatementDetailId[rowIndex].DocumentSourceAmount;
                                        newRow.DocumentSourceCurrency = oStatementDetailId[rowIndex].DocumentSourceCurrency;
                                        newRow.DocumentTreasuryCode = oStatementDetailId[rowIndex].DocumentTreasuryCode;
                                        newRow.DocumentValueDate = oStatementDetailId[rowIndex].DocumentValueDate;

                                        newRow.EntryAccountPoint = oStatementDetailId[rowIndex].EntryAccountPoint;
                                        newRow.EntryAmountBase = oStatementDetailId[rowIndex].EntryAmountBase;
                                        newRow.EntryAmountCredit = oStatementDetailId[rowIndex].EntryAmountCredit;
                                        newRow.EntryAmountCreditBase = oStatementDetailId[rowIndex].EntryAmountCreditBase;
                                        newRow.EntryAmountDebit = oStatementDetailId[rowIndex].EntryAmountDebit;
                                        newRow.EntryAmountDebitBase = oStatementDetailId[rowIndex].EntryAmountDebitBase;
                                        newRow.EntryComment = oStatementDetailId[rowIndex].EntryComment;
                                        newRow.EntryDate = oStatementDetailId[rowIndex].EntryDate;

                                        newRow.EntryDepartment = oStatementDetailId[rowIndex].EntryDepartment;
                                        newRow.EntryDocumentNumber = oStatementDetailId[rowIndex].EntryDocumentNumber;
                                        newRow.EntryId = oStatementDetailId[rowIndex].EntryId;
                                        newRow.Rate = oStatementDetailId[rowIndex].Rate;
                                        newRow.SenderDetails = oStatementDetailId[rowIndex].SenderDetails;


                                        oStatementDetail.Add(newRow);
                                    }


                                    Page++;
                                }
                            }
                        }
                    }
                    

                    if (errorText != null)
                    {
                        BDOSInternetBanking.oStatementDetailStc = null;

                        Program.uiApp.MessageBox(errorText);
                    }
                    else
                    {
                        if (oStatementDetail.Count > 0)
                        {
                            BDOSInternetBanking.fillExportMTR_BOG(form, oStatementDetail, false);
                        }
                    }
                }

                //if (operation != null)
                //{
                //    FormsB1.SimulateRefresh();
                //}

                form = null;
                operation = null;
                docEntryList = null;
            }

            catch (Exception ex)
            {
                //int errCode;
                //string errMsg;

                //Program.oCompany.GetLastError(out errCode, out errMsg);
                //errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
                Program.uiApp.SetStatusBarMessage($"{ex.Message} Inner Error: {ex.InnerException.Message}", SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            finally
            {
                GC.Collect();
            }
        }

        private static HttpClient InitializeClient()
        {
            var client = new HttpClient();

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            client.BaseAddress = new Uri(ApiBaseUrl);         
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", AuthorizeResponse.AccessToken);

            //client.SetBearerToken(AuthorizeResponse.AccessToken);
            return client;
        }

        private static void AssertToken(out string errorText)
        {
            errorText = null;
            if (AuthorizeResponse == null || AuthorizeResponse.ResponseType == AuthorizeResponse.ResponseTypes.Error)
            {
                errorText = BDOSResources.getTranslate("GetTheToken") + "!";
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                try
                {
                    SAPbouiCOM.Item oItem = oForm.Items.Item("urlWB");
                    SAPbouiCOM.ActiveX oActive = (SAPbouiCOM.ActiveX)oItem.Specific;
                    SHDocVw.WebBrowser WebBrowserChen;
                    WebBrowserChen = ((SHDocVw.WebBrowser)(oActive.Object));

                    string UrlTemp = WebBrowserChen.LocationURL;

                    if (UrlTemp.StartsWith(CallbackUri))
                    {
                        if (LocationURL == false)
                        {
                            LocationURL = true;
                            AuthorizeResponse = new AuthorizeResponse(UrlTemp);
                            //oForm.Close();
                        }
                    }
                }
                catch { }

                //if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE & pVal.BeforeAction == false)
                //{                  
                //    formClose( oForm, pVal, out errorText);
                //}

                if (pVal.ItemUID == "1" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    oForm.Close();
                    formClose(oForm, pVal, out errorText);
                }
                //if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_WEBMESSAGE)
                //{
                //    SAPbouiCOM.Item oItem = oForm.Items.Item("urlWB");
                //    SAPbouiCOM.WebBrowser oActive = (SAPbouiCOM.WebBrowser)oItem.Specific;

                //    if (pVal.BeforeAction)
                //    {
                //        string d = "d";
                //    }
                //    else
                //    {
                //        string d = "d";
                //    }
                //}
            }
        }
    }
}

