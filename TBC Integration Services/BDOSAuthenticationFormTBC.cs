using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn.TBC_Integration_Services
{
    static partial class BDOSAuthenticationFormTBC
    {
        private static string userNameField;
        private static string passwordField;
        private static string nonceField;
        private static string serviceUrlField;
        private static SAPbouiCOM.Form formField;
        private static string operationField;
        private static List<int> docEntryListField;
        private static bool importBatchPaymentOrdersField;
        private static string batchNameField;
        private static AccountMovementFilterIo oAccountMovementFilterIoField;

        public static string userName
        {
            get
            {
                return userNameField;
            }
            set
            {
                userNameField = value;
            }
        }

        public static string password
        {
            get
            {
                return passwordField;
            }
            set
            {
                passwordField = value;
            }
        }

        public static string nonce
        {
            get
            {
                return nonceField;
            }
            set
            {
                nonceField = value;
            }
        }

        public static string serviceUrl
        {
            get
            {
                return serviceUrlField;
            }
            set
            {
                serviceUrlField = value;
            }
        }

        public static SAPbouiCOM.Form form
        {
            get
            {
                return formField;
            }
            set
            {
                formField = value;
            }
        }

        public static string operation
        {
            get
            {
                return operationField;
            }
            set
            {
                operationField = value;
            }
        }

        public static List<int> docEntryList
        {
            get
            {
                return docEntryListField;
            }
            set
            {
                docEntryListField = value;
            }
        }

        public static bool importBatchPaymentOrders
        {
            get
            {
                return importBatchPaymentOrdersField;
            }
            set
            {
                importBatchPaymentOrdersField = value;
            }
        }

        public static string batchName
        {
            get
            {
                return batchNameField;
            }
            set
            {
                batchNameField = value;
            }
        }

        public static AccountMovementFilterIo oAccountMovementFilterIo
        {
            get
            {
                return oAccountMovementFilterIoField;
            }
            set
            {
                oAccountMovementFilterIoField = value;
            }
        }

        public static void createForm( SAPbouiCOM.Form formOutgoingPayment, string operationOutgoingPayment, List<int> docEntryListOutgoingPayment, bool importBatchPaymentOrdersOutgoingPayment, string batchNameOutgoingPayment, AccountMovementFilterIo oAccountMovementFilter, out string errorText)
        {
            errorText = null;
            form = formOutgoingPayment;
            operation = operationOutgoingPayment;
            docEntryList = docEntryListOutgoingPayment;
            importBatchPaymentOrders = importBatchPaymentOrdersOutgoingPayment;
            batchName = batchNameOutgoingPayment;
            oAccountMovementFilterIo = oAccountMovementFilter;

            int formHeight = Program.uiApp.Desktop.Width / 9;
            int formWidth = Program.uiApp.Desktop.Width / 4;
            
            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSAuthenticationFormTBC");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Fixed);
            formProperties.Add("Title", BDOSResources.getTranslate("LoginRequired"));
            formProperties.Add("Left", (Program.uiApp.Desktop.Width - formWidth) / 2);
            formProperties.Add("ClientWidth", formWidth);
            formProperties.Add("Top", (Program.uiApp.Desktop.Height - formHeight) / 3);
            formProperties.Add("ClientHeight", formHeight);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm( formProperties, out oForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }

            if (formExist == true)
            {
                if (newForm == true)
                {
                    errorText = null;
                    Dictionary<string, object> formItems;
                    string itemName = "";

                    int left_s = 6;
                    int left_e = 120;
                    int height = 15;
                    int top = 6;
                    int width_s = 121;
                    int width_e = 148;

                    formItems = new Dictionary<string, object>();
                    itemName = "userNameS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("UserName"));
                    formItems.Add("LinkTo", "userNameE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "userNameE"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("AffectsFormMode", false);
                    formItems.Add("Value", "B941646");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "passwordS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Password"));
                    formItems.Add("LinkTo", "PasswordE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "passwordE"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);                   
                    formItems.Add("IsPassword", true);
                    formItems.Add("AffectsFormMode", false);
                    formItems.Add("Value", "Aa123456");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "digipassS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Digipass"));
                    formItems.Add("LinkTo", "digipassE");                 

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "digipassE"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("AffectsFormMode", false);
                    formItems.Add("Value", "1111");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    //ღილაკები
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

                    left_s = left_s + width_s + 2;

                    itemName = "changePasB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s * 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("ChangePassword"));

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

        public static void createFormChangePassword( out string errorText)
        {
            errorText = null;
            int formHeight = Program.uiApp.Desktop.Width / 9;
            int formWidth = Program.uiApp.Desktop.Width / 4;
            
            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSChangePasswordFormTBC");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Fixed);
            formProperties.Add("Title", BDOSResources.getTranslate("ChangePassword"));
            formProperties.Add("Left", (Program.uiApp.Desktop.Width - formWidth) / 2);
            formProperties.Add("ClientWidth", formWidth);
            formProperties.Add("Top", (Program.uiApp.Desktop.Height - formHeight) / 3);
            formProperties.Add("ClientHeight", formHeight);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm( formProperties, out oForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }

            if (formExist == true)
            {
                if (newForm == true)
                {
                    errorText = null;
                    Dictionary<string, object> formItems;
                    string itemName = "";

                    int left_s = 6;
                    int left_e = 120;
                    int height = 15;
                    int top = 6;
                    int width_s = 121;
                    int width_e = 148;                   

                    formItems = new Dictionary<string, object>();
                    itemName = "passwordS1"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("NewPassword"));
                    formItems.Add("LinkTo", "PasswordE1");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "passwordE1"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("IsPassword", true);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "passwordS2"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("ConfirmPassword"));
                    formItems.Add("LinkTo", "PasswordE2");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "passwordE2"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("IsPassword", true);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    //ღილაკები
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

        public static void clickBDOSAuthenticationFormTBC(  SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("userNameE").Specific;
                string userNameValue = oEditText.Value;
                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("passwordE").Specific;
                string passwordValue = oEditText.Value;
                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("digipassE").Specific;
                string digipassValue = oEditText.Value;

                SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("userNameS").Specific;
                string userNameS = oStaticText.Caption;
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("passwordS").Specific;
                string passwordS = oStaticText.Caption;
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("digipassS").Specific;
                string digipassS = oStaticText.Caption;

                if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "changePasB")
                    {
                        if (string.IsNullOrEmpty(userNameValue) || string.IsNullOrEmpty(passwordValue) || string.IsNullOrEmpty(digipassValue))
                        {
                            //You can't leave empty
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + userNameS + "\", \"" + passwordS + "\", \"" + digipassS + "\"";
                            Program.uiApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
                            Program.uiApp.StatusBar.SetText(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }
                        else
                        {
                            userName = userNameValue;
                            password = passwordValue;
                            nonce = digipassValue;
                            createFormChangePassword( out errorText); //პაროლის შეცვლის ფორმა
                        }
                    }
                }
                else
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (string.IsNullOrEmpty(userNameValue) || string.IsNullOrEmpty(passwordValue) || string.IsNullOrEmpty(digipassValue))
                        {
                            //You can't leave empty
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + userNameS + "\", \"" + passwordS + "\", \"" + digipassS + "\"";
                            Program.uiApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
                            Program.uiApp.StatusBar.SetText(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);                        
                            return;
                        }
                        else
                        {
                            userName = userNameValue;
                            password = passwordValue;
                            nonce = digipassValue;
                            oForm.Close();

                            if (operation == "import") //იმპორტი
                            {
                                int answer = 1;
                                if (form != null && form.TypeEx == "426") //ე.ი დოკუმენტიდან გამოიძახება
                                {
                                    string paymentID = form.DataSources.DBDataSources.Item("OVPM").GetValue("U_paymentID", 0).Trim();

                                    if (string.IsNullOrEmpty(paymentID) == false)
                                    {
                                        answer = Program.uiApp.MessageBox("დოკუმენტი უკვე არსებობს ინტერნეტბანკში. გსურთ ხელახლა გადაგზავნა?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                                    }
                                }
                                if (answer == 2)
                                {
                                    return;
                                }

                                PaymentService oPaymentService = MainPaymentService.setPaymentService(serviceUrl, userName, password, nonce);
                                List<string> infoList = OutgoingPayment.importPaymentOrderTBC( oPaymentService, docEntryList, importBatchPaymentOrders, batchName, out errorText);
                                if (errorText != null)
                                {
                                    Program.uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    for (int i = 0; i < infoList.Count(); i++)
                                    {
                                        Program.uiApp.SetStatusBarMessage(infoList[i], SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                    }
                                    
                                    if (form != null && form.UniqueID == "BDOSInternetBankingForm") //თუ დამუშავებიდან გამოიძახება
                                    {
                                        BDOSInternetBanking.fillImportMTR( form, out errorText);
                                    }
                                }
                            }
                            else if (operation == "updateStatus") //სტატუსის განახლება
                            {
                                PaymentService oPaymentService = MainPaymentService.setPaymentService(serviceUrl, userName, password, nonce);
                                List<string> infoList = OutgoingPayment.updateStatusPaymentOrderTBC( oPaymentService, docEntryList, out errorText);
                                if (errorText != null)
                                {
                                    Program.uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    for (int i = 0; i < infoList.Count(); i++)
                                    {
                                        Program.uiApp.SetStatusBarMessage(infoList[i], SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                    }
                                    if (form != null && form.UniqueID == "BDOSInternetBankingForm") //თუ დამუშავებიდან გამოიძახება
                                    {
                                        BDOSInternetBanking.fillImportMTR( form, out errorText);
                                    }
                                }
                            }
                            else if (operation == "getData") //ჩამოტვირთვა
                            {
                                MovementService oMovementService = MainMovementService.setMovementService(serviceUrl, userName, password, nonce);
                                AccountMovementDetailIo[] oAccountMovementDetailIo = null;
                                BaseQueryResultIo oBaseQueryResultIo = MainMovementService.getAccountMovements(oMovementService, oAccountMovementFilterIo, out oAccountMovementDetailIo, out errorText);
                                
                                if (errorText != null)
                                {
                                    BDOSInternetBanking.oAccountMovementDetailIoStc = null;
                                    BDOSInternetBanking.oBaseQueryResultIoStc = null;

                                    Program.uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    BDOSInternetBanking.fillExportMTR_TBC( form, oBaseQueryResultIo, oAccountMovementDetailIo, false, out errorText);
                                }
                            }
                            if (operation != null)
                            {
                                FormsB1.SimulateRefresh();
                            }

                            form = null;
                            operation = null;
                            docEntryList = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void clickBDOSChangePasswordFormTBC(  SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("passwordE1").Specific;
                string newPassword = oEditText.Value;
                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("passwordE2").Specific;
                string confirmPassword = oEditText.Value;

                SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("passwordS1").Specific;
                string passwordS1 = oStaticText.Caption;
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("passwordS2").Specific;
                string passwordS2 = oStaticText.Caption;

                if (pVal.BeforeAction == false)
                {

                }
                else
                {
                    if (pVal.ItemUID == "1") //OK
                    {                    
                        if (string.IsNullOrEmpty(newPassword) || string.IsNullOrEmpty(confirmPassword))
                        {
                            //You can't leave empty
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + passwordS1 + "\", \"" + passwordS2 + "\"";
                            Program.uiApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
                            Program.uiApp.StatusBar.SetText(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }
                        else if (newPassword != confirmPassword)
                        {
                            errorText = BDOSResources.getTranslate("ThesePasswordsDontMatchTryAgain"); //These passwords don't match. Try again!                          
                            Program.uiApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
                            Program.uiApp.StatusBar.SetText(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }
                        else
                        {
                            string clientID;
                            int port;
                            string url = CommonFunctions.getServiceUrlForInternetBanking( "TBC", out clientID, out port, out errorText);
                            if (string.IsNullOrEmpty(errorText) == false)
                            {
                                Program.uiApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
                                Program.uiApp.StatusBar.SetText(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                return;
                            }
                            serviceUrl = url;

                            ChangePasswordService oChangePasswordService = MainChangePasswordService.setChangePasswordService(serviceUrl, userName, password, nonce);
                            string result = MainChangePasswordService.changePassword(oChangePasswordService, newPassword, out errorText);
                            if (string.IsNullOrEmpty(errorText))
                            {
                                Program.uiApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
                                Program.uiApp.StatusBar.SetText(result, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                oForm.Close();
                            }
                            else
                            {
                                Program.uiApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
                                Program.uiApp.StatusBar.SetText(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                return;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (FormUID == "BDOSAuthenticationFormTBC") //აუთენთიფიკაციის ფორმა
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                    {
                        string clientID;
                        int port;
                        string url = CommonFunctions.getServiceUrlForInternetBanking( "TBC", out clientID, out port, out errorText);

                        if (string.IsNullOrEmpty(errorText) == false)
                            {
                                Program.uiApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
                                Program.uiApp.StatusBar.SetText(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                        serviceUrl = url;
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        clickBDOSAuthenticationFormTBC( oForm, pVal, out errorText);
                    }
                }
                else if (FormUID == "BDOSChangePasswordFormTBC") //პაროლის შეცვლის ფორმა
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        clickBDOSChangePasswordFormTBC( oForm, pVal, out errorText);
                    }
                }
            }
        }
    }
}
