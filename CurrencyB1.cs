using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class CurrencyB1
    {
        public static List<string> getCurrencyList(out string errorText)
        {
            errorText = null;
            List<string> currencyListFromDB = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string MainCurncy = getMainCurrency(out errorText);
            if (errorText != null)
            {
                return currencyListFromDB;
            }

            //string query = "SELECT " + 
            //    "\"CurrCode\" " + 
            //    "FROM " + "\"" + Program.oCompany.CompanyDB + "\"" + "." + "\"OCRN\" " + 
            //    "WHERE " + "\"CurrCode\"" + " != '" + MainCurncy + "'";

            string query = @"SELECT ""CurrCode"" FROM ""OCRN"" WHERE ""CurrCode"" != '" + MainCurncy + "'";

            try
            {
                oRecordSet.DoQuery(query);
                currencyListFromDB = new List<string>();

                while (!oRecordSet.EoF)
                {
                    currencyListFromDB.Add(oRecordSet.Fields.Item("CurrCode").Value.ToString());
                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorOfCurrencyList") + " " + BDOSResources.getTranslate("ErrorDescription") + " " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "" + BDOSResources.getTranslate("OtherInfo") + ": " + ex.Message;
            }

            return currencyListFromDB;
        }

        public static string getMainCurrency(out string errorText)
        {
            errorText = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "SELECT \"MainCurncy\" , \"SysCurrncy\" FROM \"OADM\"";
            string MainCurncy = null;

            try
            {
                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    MainCurncy = oRecordSet.Fields.Item("MainCurncy").Value.ToString();
                    return MainCurncy;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }

            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }

            return MainCurncy;
        }

        public static bool importCurrencyRate(string dateStr, out string errorText, List<string> currencyList = null)
        {
            errorText = null;
            DateTime date = DateTime.TryParse(dateStr, out date) == false ? DateTime.Today : date;

            if (currencyList == null)
            {
                currencyList = getCurrencyList(out errorText);
                if (errorText != null)
                {
                    return false;
                }
            }

            NBGCurrency NBGCurrencyObj = new NBGCurrency();
            Dictionary<string, List<object>> currencyMapFromNBG = NBGCurrencyObj.GetCurrencyRateList(dateStr);

            List<object> currencyValueFromNBG = new List<object>();
            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            if (currencyMapFromNBG != null)
            {
                for (int i = 0; i < currencyList.Count; i++)
                {
                    if (currencyMapFromNBG.TryGetValue(CommonFunctions.getCurrencyInternationalCode(currencyList[i]), out currencyValueFromNBG))
                    {
                        oSBOBob.SetCurrencyRate(currencyList[i], date, Convert.ToDouble(currencyValueFromNBG[1], CultureInfo.InvariantCulture), true);
                    }
                }
            }

            Marshal.ReleaseComObject(oSBOBob);

            return true;
        }

        public static bool importCurrencyRate(string dateStr, out string errorText, ref Dictionary<string, Dictionary<int, double>> currencyListFromNBG, List<string> currencyList = null)
        {
            errorText = null;
            DateTime date = DateTime.TryParse(dateStr, out date) == false ? DateTime.Today : date;

            if (currencyList == null)
            {
                currencyList = getCurrencyList(out errorText);
                if (errorText != null)
                {
                    return false;
                }
            }

            NBGCurrency NBGCurrencyObj = new NBGCurrency();
            Dictionary<string, List<object>> currencyMapFromNBG = NBGCurrencyObj.GetCurrencyRateList(dateStr);

            List<object> currencyValueFromNBG = new List<object>();
            Dictionary<int, double> dailyRate;
            string currencyInternationalCode;

            if (currencyMapFromNBG != null)
            {
                for (int i = 0; i < currencyList.Count; i++)
                {
                    currencyInternationalCode = CommonFunctions.getCurrencyInternationalCode(currencyList[i]);

                    if (currencyMapFromNBG.TryGetValue(currencyInternationalCode, out currencyValueFromNBG))
                    {
                        dailyRate = new Dictionary<int, double>();
                        dailyRate.Add(date.Day, Convert.ToDouble(currencyValueFromNBG[1], CultureInfo.InvariantCulture));
                        if (currencyListFromNBG.Keys.Contains(currencyInternationalCode))
                        {
                            currencyListFromNBG[currencyInternationalCode].Add(date.Day, Convert.ToDouble(currencyValueFromNBG[1], CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            currencyListFromNBG.Add(currencyInternationalCode, dailyRate);
                        }
                    }
                }
            }

            return true;
        }
    }
}
