using BDO_Localisation_AddOn.BOG_Integration_Services.Model;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn.BOG_Integration_Services
{
    static partial class BDOSAuthenticationFormBOG
    {
        private static SAPbouiCOM.Form _form;
        private static string _operation;
        private static List<int> _docEntryList;
        private static bool _importBatchPaymentOrders;
        private static string _batchName;
        private static string _authorizeUri; // BOG
        private static string _callbackUri;
        private static AuthorizeResponse _authorizeResponse;
        private static string _apiBaseUrl;
        private static StatementFilter _oStatementFilter; // BOG

        public static void createForm(SAPbouiCOM.Form formOutgoingPayment, string operationOutgoingPayment, List<int> docEntryListOutgoingPayment, bool importBatchPaymentOrdersOutgoingPayment, string batchNameOutgoingPayment, StatementFilter oStatementFilterTmp, out string errorText)
        {
            //---> BOG
            var wsdl = CommonFunctions.getServiceWSDLForInternetBanking("BOG", out string clientId, out int port, out string mode, out string url, out errorText);
            var scope = "read write";

            if (!string.IsNullOrEmpty(errorText))
                return;

            InitializeGlobalVariables();

            string authorizeUrl = GetAuthorizeUrl(scope, clientId);
            //---> BOG

            var authForm = GetAuthForm();

            if (authForm == string.Empty)
            {
                errorText = "Auth form is not chosen";
                return;
            }

            if (authForm == "Explorer")
                HandleInternetExplorerCase(authorizeUrl);
            else
                HandleChromeCase(authorizeUrl);

            GC.Collect();

            void InitializeGlobalVariables()
            {
                _form = formOutgoingPayment;
                _operation = operationOutgoingPayment;
                _docEntryList = docEntryListOutgoingPayment;
                _importBatchPaymentOrders = importBatchPaymentOrdersOutgoingPayment;
                _batchName = batchNameOutgoingPayment;
                _oStatementFilter = oStatementFilterTmp;
                _apiBaseUrl = wsdl;

                if (mode == "test") //სატესტო გარემოში port შევსებული უნდა იყოს
                {
                    _authorizeUri = $"{url}:{port}/Oauth/Connect/Authorize.aspx";
                    _callbackUri = $"{url}:{port}/Oauth/Connect/Token.aspx";
                }
                else if (mode == "real")
                {
                    _authorizeUri = $"{url}/Oauth/Connect/Authorize.aspx";
                    _callbackUri = $"{url}/Oauth/Connect/Token.aspx";
                }
                else if (mode == "realNew")
                {
                    _authorizeUri = $"{url}/auth/realms/bog/protocol/openid-connect/auth";
                    _callbackUri = $"{url}/auth/realms/bog/protocol/openid-connect/token";
                    scope = "corp";
                }
                else if (mode == "testNew")
                {
                    _authorizeUri = $"{url}/auth/realms/bog-test/protocol/openid-connect/auth";
                    _callbackUri = $"{url}/auth/realms/bog-test/protocol/openid-connect/token";
                    scope = "corp";
                }
            }
        }

        private static string GetAuthForm()
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            var getAuthFormQuery = @"SELECT ""U_AuthForm""
                                        FROM ""@BDO_INTB""
                                        WHERE ""@BDO_INTB"".""U_program"" = 'BOG'";

            oRecordSet.DoQuery(getAuthFormQuery);

            if (oRecordSet.RecordCount > 0)
            {
                return oRecordSet.Fields.Item("U_AuthForm").Value;
            }
            else
            {
                return string.Empty;
            }
        }

        private static void HandleChromeCase(string authorizeUrl)
        {
            var chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;

            using (var controller = new WebDriverController())
            {
                controller.Driver = new ChromeDriver(chromeDriverService, new ChromeOptions())
                {
                    Url = authorizeUrl
                };

                while (true)
                {
                    if (controller.Driver == null)
                        return;

                    string url = controller.Driver.Url;
                    _authorizeResponse = GetAuthorizeResponse(url);

                    if (_authorizeResponse != null)
                        break;
                }
            }
            HandleResponse();
        }

        private static void HandleInternetExplorerCase(string authorizeUrl)
        {
            int formHeight = Program.uiApp.Desktop.Width;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>
            {
                { "UniqueID", "BDOSAuthenticationFormBOG" },
                { "BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable },
                { "Title", BDOSResources.getTranslate("LoginRequired") },
                { "ClientWidth", formWidth },
                { "ClientHeight", formHeight }
            };

            //formProperties.Add("Left", (Program.uiApp.Desktop.Width - formWidth) / 2);
            //formProperties.Add("Top", (Program.uiApp.Desktop.Height - formHeight) / 3);

            bool formExist = FormsB1.createForm(formProperties, out SAPbouiCOM.Form oForm, out bool newForm, out string errorText);

            if (errorText != null)
                return;

            if (formExist)
            {
                if (newForm)
                {
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
                    WebBrowserChen.Navigate2(authorizeUrl);

                    Dictionary<string, object> formItems;
                    var left_s = 6;
                    var height = 19;
                    var top = oForm.ClientHeight - 25;
                    var width_s = 65;
                    string itemName = "1";

                    formItems = new Dictionary<string, object>
                    {
                        { "Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON },
                        { "Left", left_s },
                        { "Width", width_s },
                        { "Top", top },
                        { "Height", height },
                        { "UID", itemName }
                    };

                    FormsB1.createFormItem(oForm, formItems, out errorText);

                    if (errorText != null)
                        return;

                    left_s = left_s + width_s + 2;

                    itemName = "2";

                    formItems = new Dictionary<string, object>
                    {
                        { "Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON },
                        { "Left", left_s },
                        { "Width", width_s },
                        { "Top", top },
                        { "Height", height },
                        { "UID", itemName }
                    };

                    FormsB1.createFormItem(oForm, formItems, out errorText);

                    if (errorText != null)
                        return;
                }

                oForm.Visible = true;
                oForm.Select();
            }
        }

        private static string GetAuthorizeUrl(string scope, string clientId)
        {
            var client = new OAuth2Client(new Uri(_authorizeUri));
            var state = Guid.NewGuid().ToString();
            var url = client.CreateAuthorizeUrl(clientId, "token", scope, _callbackUri, state);
            return url;
        }

        private static void HandleResponse()
        {
            string errorText = null;

            try
            {
                if (_operation == "import" || _operation == "updateStatus") //იმპორტი
                {
                    AssertToken(out errorText);

                    if (string.IsNullOrEmpty(errorText) == false)
                    {
                        Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                    else
                    {
                        HttpClient client = GetHttpClient();

                        var infoList = _operation == "import"
                            ? OutgoingPayment.importPaymentOrderBOG(client, _docEntryList, _importBatchPaymentOrders, _batchName, out errorText)
                            : OutgoingPayment.updateStatusPaymentOrderBOG(client, _docEntryList, out errorText);

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

                            if (_form != null && _form.UniqueID == "BDOSInternetBankingForm") //თუ დამუშავებიდან გამოიძახება
                            {
                                BDOSInternetBanking.fillImportMTR(_form, out errorText);
                            }
                        }
                    }
                }

                else if (_operation == "getData") //ამონაწერი
                {
                    HttpClient client = GetHttpClient();
                    Task<Statement> oStatement = null;

                    List<StatementDetail> oStatementDetail = null;
                    Task<List<StatementDetail>> oStatementDetailIdTask = null;
                    List<StatementDetail> oStatementDetailId = null;

                    if (_oStatementFilter.Page == 0)
                    {
                        oStatement = MainPaymentServiceBOG.getStatement(client, _oStatementFilter.AccountNumber, _oStatementFilter.Currency, _oStatementFilter.PeriodFrom, _oStatementFilter.PeriodTo);
                        if (oStatement != null)
                        {
                            oStatementDetail = oStatement.Result.Records;

                            int id = Convert.ToInt32(oStatement.Result.Id);
                            int count = oStatement.Result.Count;

                            if (count > 1000)
                            {
                                double count10000 = count / 1000;
                                double Page = 2;
                                while (Page <= count10000 + 1)
                                {
                                    oStatementDetailIdTask = MainPaymentServiceBOG.getStatementByID(client, _oStatementFilter.AccountNumber, _oStatementFilter.Currency, id, Convert.ToInt32(Page));
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
                            BDOSInternetBanking.fillExportMTR_BOG(_form, oStatementDetail, false);
                        }
                    }
                }

                _form = null;
                _operation = null;
                _docEntryList = null;
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

        private static HttpClient GetHttpClient()
        {
            var client = new HttpClient();

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            client.BaseAddress = new Uri(_apiBaseUrl);
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _authorizeResponse.AccessToken);

            //client.SetBearerToken(AuthorizeResponse.AccessToken);
            return client;
        }

        private static void AssertToken(out string errorText)
        {
            errorText = null;

            if (_authorizeResponse == null || _authorizeResponse.ResponseType == AuthorizeResponse.ResponseTypes.Error)
            {
                errorText = BDOSResources.getTranslate("GetTheToken") + "!";
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                return;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

            try
            {
                SAPbouiCOM.Item oItem = oForm.Items.Item("urlWB");
                SAPbouiCOM.ActiveX oActive = (SAPbouiCOM.ActiveX)oItem.Specific;
                SHDocVw.WebBrowser WebBrowserChen;
                WebBrowserChen = (SHDocVw.WebBrowser)(oActive.Object);

                _authorizeResponse = GetAuthorizeResponse(WebBrowserChen.LocationURL);
            }
            catch { }

            if (pVal.ItemUID == "1" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
            {
                oForm.Close();
                HandleResponse();
            }

            #region Commented section
            //if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE & pVal.BeforeAction == false)
            //{
            //    formClose( oForm, pVal, out errorText);
            //}

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
            #endregion
        }

        private static AuthorizeResponse GetAuthorizeResponse(string url)
        {
            if (url.StartsWith(_callbackUri))
            {
                return _authorizeResponse ?? new AuthorizeResponse(url);
            }

            return null;
        }

        private class WebDriverController : IDisposable
        {
            public IWebDriver Driver { get; set; }

            public void Dispose()
            {
                this.Driver.Dispose();
            }
        }
    }
}