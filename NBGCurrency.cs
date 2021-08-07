using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;

namespace BDO_Localisation_AddOn
{
    class NBGCurrency
    {
        private NBGCurrencyService_HTTP.NBGCurrencyService NBG_CurrencyPortClient_field = null;

        public NBGCurrency()
        {
            this.NBG_CurrencyPortClient_field = new NBGCurrencyService_HTTP.NBGCurrencyService();
        }

        public NBGCurrencyService_HTTP.NBGCurrencyService NBG_CurrencyPortClient
        {
            get
            {
                return this.NBG_CurrencyPortClient_field;
            }
            set
            {
                this.NBG_CurrencyPortClient_field = value;
            }
        }

        /// <summary>ვალუტის კურსის მიღება</summary>
        /// <param name="currency"></param>
        /// <param name="errorText"></param>
        /// <returns>აბრუნებს ვალუტის კურსს</returns>
        public string GetCurrency(Currency currency)
        {
            string getCurrencyResult = null;

            try
            {
                getCurrencyResult = NBG_CurrencyPortClient.GetCurrency(currency.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return getCurrencyResult;
        }

        /// <summary>ვალუტის აღწერის მიღება</summary>
        /// <param name="currency"></param>
        /// <param name="errorText"></param>
        /// <returns>აბრუნებს ვალუტის აღწერას</returns>
        public string GetCurrencyDescription(Currency currency)
        {
            string getCurrencyDescriptionResult = null;

            try
            {
                getCurrencyDescriptionResult = NBG_CurrencyPortClient.GetCurrencyDescription(currency.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return getCurrencyDescriptionResult;
        }

        /// <summary>ვალუტის ცვლილების მნიშვნელობის მიღება</summary>
        /// <param name="currency"></param>
        /// <param name="errorText"></param>
        /// <returns>აბრუნებს ვალუტის ცვლილების მნიშვნელობას</returns>
        public string GetCurrencyChange(Currency currency)
        {
            string getCurrencyChangeResult = null;

            try
            {
                getCurrencyChangeResult = NBG_CurrencyPortClient.GetCurrencyChange(currency.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return getCurrencyChangeResult;

        }

        /// <summary>ვალუტის ცვლილება მოხდა თუ არა</summary>
        /// <param name="currency"></param>
        /// <param name="errorText"></param>
        /// <returns>1 - თუ გაიზარდა; -1 - თუ დაიკლო, 0 - თუ იგივე დარჩა</returns>
        public int GetCurrencyRate(Currency currency)
        {
            int getCurrencyRateResult = 0;

            try
            {
                getCurrencyRateResult = NBG_CurrencyPortClient.GetCurrencyRate(currency.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return getCurrencyRateResult;

        }

        /// <summary>კურსების შესაბამის თარიღის მიღება</summary>
        /// <param name="errorText"></param>
        /// <returns>აბრუნებს კურსების შესაბამის თარიღს</returns>
        public string GetDate()
        {
            string getDateResult = null;

            try
            {
                getDateResult = NBG_CurrencyPortClient.GetDate();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return getDateResult;
        }

        /// <summary>კურსების ჩამოტვირთვა ყველა ვალუტისთვის</summary>
        /// <param name="errorText"></param>
        /// <returns>აბრუნებს კურსებს ყველა ვალუტისთვის</returns>
        public Dictionary<string, List<object>> GetCurrencyRateList(string dateStr)
        {
            XmlDocument xDoc = null;
            Dictionary<string, List<object>> currencyMap = null;

            try
            {
                xDoc = new XmlDocument();
                xDoc.Load("http://www.nbg.gov.ge/rss.php?date=" + dateStr);
            }

            catch (Exception ex)
            {
                throw new Exception(BDOSResources.getTranslate("ErrorWhileImportRateServiceCall") + " ERROR : " + ex.Message);
            }

            string valueXML = null;

            XmlNode currencyListNode = xDoc.SelectSingleNode("/rss/channel/item/description");

            valueXML = currencyListNode.InnerText;
            string valueStr = null;

            if (valueXML.Contains("gif" + '"'.ToString() + "/>"))
            {
                valueStr = valueXML;
            }
            else
            {
                valueStr = valueXML.Replace("gif" + '"'.ToString() + ">", "gif" + '"'.ToString() + "/>");
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(string.Format("<root>{0}</root>", valueStr));

            currencyMap = new Dictionary<string, List<object>>();

            foreach (XmlElement node in xmlDoc.GetElementsByTagName("tr"))
            {
                XmlNodeList childNodes = node.ChildNodes;
                string currency = childNodes[0].InnerText;
                string currencyDescription = childNodes[1].InnerText;
                int indexOf = currencyDescription.IndexOf(" ");
                string divider = currencyDescription.Substring(0, indexOf);
                double currencyRate = Convert.ToDouble(childNodes[2].InnerText, CultureInfo.InvariantCulture);
                currencyRate = currencyRate / Convert.ToDouble(divider, CultureInfo.InvariantCulture);
                string currencyGif = childNodes[3].ChildNodes[0].Attributes[0].Value;
                double currencyChange = Convert.ToDouble(childNodes[4].InnerText, CultureInfo.InvariantCulture) * ((currencyGif.Contains("red")) ? (-1) : 1);

                currencyMap.Add(currency, new List<object>() { currencyDescription, currencyRate, currencyGif, currencyChange });
            }

            return currencyMap;
        }

        public List<NBGCurrencyModel> GetCurrencyRateList()
        {
            try
            {
                List<NBGCurrencyModel> currencies = null;

                using (WebClient wc = new WebClient())
                {
                    var jsonString = wc.DownloadString("https://nbg.gov.ge/gw/api/ct/monetarypolicy/currencies/ka/json");
                    var array = JArray.Parse(jsonString);
                    foreach (var item in array.Children().Children())
                    {
                        if (item.Next == null)
                        {
                            var a = item.Children().Children().ToList();

                            currencies = a.Select(p => new NBGCurrencyModel
                            {
                                Code = (string)p["code"],
                                Quantity = (int)p["quantity"],
                                RateFormated = (double)p["rateFormated"],
                                DiffFormated = (double)p["diffFormated"],
                                Rate = (double)p["rate"],
                                Name = (string)p["name"],
                                Diff = (double)p["diff"],
                                Date = (DateTime)p["date"],
                                ValidFromDate = (DateTime)p["validFromDate"]
                            }).ToList();
                        }
                    }
                }
                return currencies;
            }
            catch
            {
                throw;
            }
        }
    }

    class NBGCurrencyModel
    {
        public string Code { get; set; }
        public int Quantity { get; set; }
        public double RateFormated { get; set; }
        public double DiffFormated { get; set; }
        public double Rate { get; set; }
        public string Name { get; set; }
        public double Diff { get; set; }
        public DateTime Date { get; set; }
        public DateTime ValidFromDate { get; set; }
    }

    public enum Currency
    {
        AED, AMD, AUD, AZN, BGN, BYR, CAD, CHF, CNY, CZK, DKK, EEK, EGP, EUR, GBP, HKD, HUF, ILS, INR, IRR, ISK, JPY, KGS, KWD, KZT, LTL, LVL, MDL, NOK, NZD, PLN, RON, RSD, RUB, SEK, SGD, TJS, TMT, TRY, UAH, USD, UZS
    };
}
