using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Data;
using System.IO;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Net;
using System.Runtime.InteropServices;

namespace BDO_Localisation_AddOn
{

    class License
    {
        private LicenseService.ПроверкаОбновленииКонфигурации oCheck_conf_update = null;

        public License()
        {
            this.oCheck_conf_update = new LicenseService.ПроверкаОбновленииКонфигурации();
            oCheck_conf_update.Credentials = new NetworkCredential("WebUser", "17D$9sa");
        }

        public string LicenseCompany(string encryptedText)
        {
            string result = oCheck_conf_update.ЗалицензироватьКонфигурацию(encryptedText);
            return result;
        }

        public bool LicenseSuccessfull(string licenseKey)
        {
            string stringSend = "";
            stringSend = stringSend + "<СерииныйНомер>" + licenseKey + "</СерииныйНомер>" + Environment.NewLine;
            stringSend = stringSend + "<СерииныйНомерДоп>" + "" + "</СерииныйНомерДоп>" + Environment.NewLine;
            stringSend = stringSend + "<ДанныеВебКлиента>" + "" + "</ДанныеВебКлиента>" + Environment.NewLine;
            stringSend = stringSend + "<Конфигурация>" + "BDO Localisation" + "</Конфигурация>" + Environment.NewLine;
            stringSend = stringSend + "<ВерсияКонфигурации>" + "1.0.0.0" + "</ВерсияКонфигурации>" + Environment.NewLine;

            Dictionary<string, string> CompanyInfo = CommonFunctions.getCompanyInfo();
            stringSend = stringSend + "<ИННКомпании>" + CompanyInfo["FreeZoneNo"] + "</ИННКомпании>" + Environment.NewLine;
            stringSend = stringSend + "<НаименованиеКомпании>" + CompanyInfo["CompnyName"] + "</НаименованиеКомпании>" + Environment.NewLine;
            stringSend = stringSend + "<ЮрАдресКомпании>" + CompanyInfo["CompnyAddr"] + "</ЮрАдресКомпании>" + Environment.NewLine;
            stringSend = stringSend + "<ФизАдресКомпании>" + CompanyInfo["CompnyAddr"] + "</ФизАдресКомпании>" + Environment.NewLine;

            int employeeQuantity = 0;
            stringSend = stringSend + "<КоличествоПользователей>" + employeeQuantity + "</КоличествоПользователей>" + Environment.NewLine;

            string dataBaseID = "";
            stringSend = stringSend + "<ИдентификаторБазы>" + dataBaseID + "</ИдентификаторБазы>";

            Random rnd = new Random();
            int rndKey = rnd.Next(48, 122);

            string encryptedText = CryptText(stringSend, rndKey) + Convert.ToChar(rndKey).ToString();

            string resultWebService = "";
            try
            {
                resultWebService = LicenseCompany(encryptedText);
            }
            catch { }

            if (resultWebService != "")
            {
                string deCryptText = CryptText(resultWebService, 0);

                string LicenseStatus = GetValueTeg("ЛицензияАктивна", deCryptText);
                if (LicenseStatus == BDOSResources.getTranslate("Active"))
                {
                    return UpdateLicenseInfo_OADM(resultWebService);
                }
            }
            else if (licenseKey == "")
            {
                return UpdateLicenseInfo_OADM(resultWebService);
            }

            return false;
        }

        public bool UpdateLicenseInfo_OADM(string licenseKeyResult)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            bool result = false;

            try
            {
                string query = @"UPDATE ""OADM"" SET ""OADM"".""U_BDOSLocLic"" = N'" + licenseKeyResult + "'";
                
                oRecordSet.DoQuery(query);
                result = true;
            }
            catch (Exception ex)
            {
                string errorText = ex.Message;
            }            
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;                
            }

            return result;
        }

        public string GetValueTeg(string nameTeg, string Text)
        {

            int begin = Text.IndexOf("<" + nameTeg + ">");
            begin = begin + 2 + nameTeg.Count();

            int end = Text.IndexOf("</" + nameTeg + ">");
            int quantSymbol = end - begin;

            if (end == 0)
            {
                if (nameTeg == "КоличествоЛицензииДоп" || nameTeg == "КоличествоЛицензии")
                {
                    return "0";
                }

            }

            string result = "";
            try
            {
                result = Text.Substring(begin, quantSymbol);
            }
            catch { }

            if (nameTeg == "ЛицензияАктивна")
            {
                if (result == "ИСТИНА")
                {
                    result = BDOSResources.getTranslate("Active");
                }
                else
                {
                    result = BDOSResources.getTranslate("Inactive");

                }
            }

            return result;
        }

        public string CryptText(string text, int rndKey)
        {
            int key = 0;
            if (rndKey != 0)
            {
                key = rndKey;
            }
            else
            {
                char symbolKey = text.Last();
                key = Convert.ToInt32(symbolKey);
                text = text.Remove(text.Count() - 1);
            }

            string keyText = GetKey(rndKey);

            int keyTextLength = keyText.Count();

            decimal decKeyTr = key / keyTextLength;
            int newKey = key - (int)(keyTextLength * Math.Truncate(decKeyTr));

            string symbol = "";
            int idSymbol = 0;
            string newText = "";
            string newSymbol;

            int textLength = text.Count();

            for (int i = 0; i < textLength; i++)
            {
                symbol = text.Substring(i, 1);
                idSymbol = keyText.IndexOf(symbol);

                if (idSymbol == -1)
                {
                    newSymbol = symbol;
                }
                else
                {
                    int newIDSymbol = idSymbol + newKey;
                    decimal decTr = (decimal)newIDSymbol / keyTextLength;
                    newIDSymbol = newIDSymbol - (int)(keyTextLength * Math.Truncate(decTr));
                    newSymbol = keyText.Substring(newIDSymbol, 1);
                }

                newText = newText + newSymbol;

            }

            return newText;

        }

        private string GetKey(int rndKey)
        {
            string keyText = "";

            if (rndKey != 0)
            {
                keyText = "ხ9-ЬUciჯoჭHqfИჩs,Еmзл*+მ>იFМლмhЙЩVქNვ8цზEPДп4дკსოСКтАВMjvБк<gШврჟуძРй0эЗSDk5оYЯПTBеЛбЮь(:ყბუ= @GKნbw1ფlФშ7ЦОzxa;გp_6ЭCЪгичWфeшхჰთнrХღж).НtRQტnცщЫаეწ2ГარыЧJXOAIсპ?dLZУя/yТЖuю3დъ";
            }
            else
            {
                keyText = "ъდ3юuЖТy/яУZLd?პсIAOXJЧыრაГ2წეаЫщცnტQRtН.)жღХrнთჰхшeфWчигЪCЭ6_pგ;axzОЦ7შФlფ1wbნKG@ =უბყ:(ьЮбЛеBTПЯYо5kDSЗэ0йРძуჟрвШg<кБvjMВАтКСოსკд4пДPEზц8ვNქVЩЙhмლМFი>მ+*лзmЕ,sჩИfqHჭoჯicUЬ-9ხ";
            }

            return keyText;
        }

        public static void createUserFields(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSLocLic");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Localisation Licensing Data");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Memo);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void UpdateAddOnLicense()
        {
            Dictionary<string, string> CompanyLicenseInfo = CommonFunctions.getCompanyLicenseInfo();

            if (string.IsNullOrEmpty(CompanyLicenseInfo["LicenseKey"]))
            {
                string errorText;
                createAddOnLicenseForm(out errorText);
            }
            else
            {
                string licenseUpdateDate = CompanyLicenseInfo["LicenseUpdateDate"];
                if (licenseUpdateDate != DateTime.Today.ToString("dd.MM.yyyy"))
                {
                    string licenseKey = CompanyLicenseInfo["LicenseKey"];
                    License oLicense = new License();
                    bool result = oLicense.LicenseSuccessfull(licenseKey);
                }
            }
        }

        public static void createAddOnLicenseForm(out string errorText)
        {
            errorText = null;

            Dictionary<string, string> CompanyLicenseInfo = CommonFunctions.getCompanyLicenseInfo();
            string licenseKey = CompanyLicenseInfo["LicenseKey"];

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSLocLicForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Fixed);
            formProperties.Add("Title", BDOSResources.getTranslate("LocalisationLicense"));
            formProperties.Add("Left", 500);
            formProperties.Add("ClientWidth", 250);
            formProperties.Add("Top", 500);
            formProperties.Add("ClientHeight", 35);

            SAPbouiCOM.Form oSetPasForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oSetPasForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }
            if (formExist == true)
            {
                if (newForm == true)
                {
                    Dictionary<string, object> formItems;
                    string itemName;
                    int top = 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "LicKey";
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", 7);
                    formItems.Add("Width", 125);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 14);
                    formItems.Add("Caption", BDOSResources.getTranslate("LicenseKey"));
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "BDOSLicKey");
                    formItems.Add("RightJustified", false);

                    FormsB1.createFormItem(oSetPasForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "BDOSLicKey";
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", 130);
                    formItems.Add("Width", 163);
                    formItems.Add("Top", top + 1);
                    formItems.Add("Height", 14);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("IsPassword", false);
                    formItems.Add("Description", BDOSResources.getTranslate("LicenseKey"));
                    formItems.Add("RightJustified", false);
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 150);
                    formItems.Add("Value", licenseKey);

                    FormsB1.createFormItem(oSetPasForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + 15;

                    top = top + 20;

                    itemName = "3";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", 7);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Update"));

                    FormsB1.createFormItem(oSetPasForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                }

                oSetPasForm.Visible = true;
                oSetPasForm.Select();
            }

            GC.Collect();
        }

    }
}
