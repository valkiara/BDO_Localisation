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
        public static List<string> getCurrencyList()
        {
            List<string> currencyListFromDB = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string MainCurncy = getMainCurrency(out var errorText);
            if (errorText != null)
            {
                return currencyListFromDB;
            }

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
                return currencyListFromDB;
            }
            catch
            {
                throw;
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordSet);
            }
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

        public static void importCurrencyRate(string dateStr, List<string> currencyList = null)
        {
            DateTime date = DateTime.TryParse(dateStr, out date) == false ? DateTime.Today : date;

            if (currencyList == null)
                currencyList = getCurrencyList();

            NBGCurrency NBGCurrencyObj = new NBGCurrency();
            List<NBGCurrencyModelXml> currencyMapFromNBG = NBGCurrencyObj.GetCurrencyRateList(date);

            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            if (currencyMapFromNBG != null && currencyMapFromNBG.Count>0)
            {
                for (int i = 0; i < currencyList.Count; i++)
                {
                    double rate = currencyMapFromNBG.Where(x => x.Currency == CommonFunctions.getCurrencyInternationalCode(currencyList[i])).Select(x => x.CurrencyRate).FirstOrDefault();

                    if (rate != 0.0)
                        oSBOBob.SetCurrencyRate(currencyList[i], date, rate, true);
                }
            }

            Marshal.ReleaseComObject(oSBOBob);
        }

        public static void importCurrencyRate(string dateStr, ref Dictionary<string, Dictionary<int, double>> currencyListFromNBG, List<string> currencyList = null)
        {
            DateTime date = DateTime.TryParse(dateStr, out date) == false ? DateTime.Today : date;

            if (currencyList == null)
                currencyList = getCurrencyList();

            NBGCurrency NBGCurrencyObj = new NBGCurrency();
            List<NBGCurrencyModelXml> currencyMapFromNBG = NBGCurrencyObj.GetCurrencyRateList(date);

            Dictionary<int, double> dailyRate;
            string currencyInternationalCode;

            if (currencyMapFromNBG != null && currencyMapFromNBG.Count > 0)
            {
                for (int i = 0; i < currencyList.Count; i++)
                {
                    currencyInternationalCode = CommonFunctions.getCurrencyInternationalCode(currencyList[i]);

                    double rate = currencyMapFromNBG.Where(x => x.Currency == currencyInternationalCode).Select(x => x.CurrencyRate).FirstOrDefault();

                    if (rate != 0.0)
                    {
                        dailyRate = new Dictionary<int, double>();
                        dailyRate.Add(date.Day, rate);
                        if (currencyListFromNBG.Keys.Contains(currencyInternationalCode))
                            currencyListFromNBG[currencyInternationalCode].Add(date.Day, rate);
                        else
                            currencyListFromNBG.Add(currencyInternationalCode, dailyRate);
                    }
                }
            }
        }
    }
}
