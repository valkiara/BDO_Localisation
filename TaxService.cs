using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BDO_Localisation_AddOn
{
    class TaxService
    {
        private TaxService_HTTP.Service taxService_soapClient_field;
        private string login_field;
        private string password_field;

        public TaxService()
        {
            this.taxService_soapClient_field = new TaxService_HTTP.Service();
            this.login_field = "ibs";
            this.password_field = "i93b46s";
        }

        public TaxService(string login, string password)
        {
            this.taxService_soapClient_field = new TaxService_HTTP.Service();
            this.login_field = login;
            this.password_field = password;
        }

        public string login
        {
            get
            {
                return this.login_field;
            }
            set
            {
                this.login_field = value;
            }
        }

        public string password
        {
            get
            {
                return this.password_field;
            }
            set
            {
                this.password_field = value;
            }
        }

        public Dictionary<string, object> GetTPInfo(string TIN, out string errorText)
        {
            errorText = null;

            TaxService_HTTP.ExtendedTPInfo oExtendedTPInfo;
            Dictionary<string, object> GetTPInfo_result = null;

            try
            {
                oExtendedTPInfo = taxService_soapClient_field.GetTPInfo(login, password, TIN);

                //oExtendedTPInfo = taxService_soapClient_field.GetTPInfo(login, password, "01017027734");//ინდ.მეწარმე
                //oExtendedTPInfo = taxService_soapClient_field.GetTPInfo(login, password, "406118320");//სპს
                //oExtendedTPInfo = taxService_soapClient_field.GetTPInfo(login, password, "441485251");//კოოპერატივი
                //oExtendedTPInfo = taxService_soapClient_field.GetTPInfo(login, password, "400185731");//შპს
                //oExtendedTPInfo = taxService_soapClient_field.GetTPInfo(login, password, "404519865");//სააქციო საზოგადოება
                //oExtendedTPInfo = taxService_soapClient_field.GetTPInfo(login, password, "216307912");//კომანდიტური საზოგადოება
                //oExtendedTPInfo = taxService_soapClient_field.GetTPInfo(login, password, "416329066");//არასამეწარმეო (არაკომეციული) იურ.პირი
                //oExtendedTPInfo = taxService_soapClient_field.GetTPInfo(login, password, "406164234");//საჯარო სამართლის იურიდიული პირი
                //oExtendedTPInfo = taxService_soapClient_field.GetTPInfo(login, password, "404520327");//უცხოური საწარმოს ფილიალი
                //oExtendedTPInfo = taxService_soapClient_field.GetTPInfo(login, password, "405131959");//უცხოური არასამეწარმეო იურ.პირის ფილიალი
                //oExtendedTPInfo = taxService_soapClient_field.GetTPInfo(login, password, "206322102"); //ამხანაგობა
                //oExtendedTPInfo = taxService_soapClient_field.GetTPInfo(login, password, "12345678910");//ფიზიკური პირი

                if (oExtendedTPInfo != null)
                {
                    GetTPInfo_result = new Dictionary<string, object>();
                    GetTPInfo_result.Add("ActivityType", oExtendedTPInfo.ActivityType);
                    GetTPInfo_result.Add("Address", oExtendedTPInfo.Address);
                    GetTPInfo_result.Add("LastRegistrationChange", oExtendedTPInfo.LastRegistrationChange);
                    GetTPInfo_result.Add("Name", oExtendedTPInfo.Name);
                    GetTPInfo_result.Add("Notes", oExtendedTPInfo.Notes);
                    GetTPInfo_result.Add("OrganizationType", oExtendedTPInfo.OrganizationType);
                    GetTPInfo_result.Add("OrganizationTypeShort", oExtendedTPInfo.OrganizationTypeShort);
                    GetTPInfo_result.Add("PersonId", oExtendedTPInfo.PersonId);
                    GetTPInfo_result.Add("RegistrationDate", oExtendedTPInfo.RegistrationDate);
                    GetTPInfo_result.Add("RegistrationNumber", oExtendedTPInfo.RegistrationNumber);
                    GetTPInfo_result.Add("Representatives", oExtendedTPInfo.Representatives);
                    GetTPInfo_result.Add("SaIdentNo", oExtendedTPInfo.SaIdentNo);
                    GetTPInfo_result.Add("Status", oExtendedTPInfo.Status);
                }
            }

            catch (Exception ex)
            {
                errorText = "შეცდომა მოხდა სამეწარმეო რეესტრის სერვისის გამოძახებისას! GetTPInfo, ERROR : " + ex.Message;
                return GetTPInfo_result;
            }

            return GetTPInfo_result;
        }
    }
}
