using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    static partial class BDOSAutomaticTasks
    {
        public static void importCurrencyRate()
        {
            DateTime startDate = DateTime.Today;
            DateTime endDate = DateTime.Today;
            string startDateStr;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT TOP 1 * FROM ""ORTT"" WHERE ""RateDate""<='" + DateTime.Today.ToString("yyyyMMdd") + @"' ORDER BY ""RateDate""  DESC";

            try
            {
                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    startDate = DateTime.TryParse(oRecordSet.Fields.Item("RateDate").Value.ToString("yyyyMMdd") == "18991230" ? DateTime.Today : oRecordSet.Fields.Item("RateDate").Value.ToString(), out startDate) == false ? DateTime.Today : startDate;
                }

                if (startDate != endDate)
                {
                    Program.uiApp.SetStatusBarMessage($"{BDOSResources.getTranslate("ImportingCurrencies")}...", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                    startDate = startDate.AddDays(1);

                    while (startDate <= endDate)
                    {
                        startDateStr = startDate.ToString("yyyy-MM-dd");
                        CurrencyB1.importCurrencyRate(startDateStr);
                        startDate = startDate.AddDays(1);
                    }
                    Program.uiApp.SetStatusBarMessage($"{BDOSResources.getTranslate("CurrenciesHaveBeenImportedSuccessfully")}!", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
            }
            catch (Exception ex)
            {
                Program.uiApp.SetStatusBarMessage($"{BDOSResources.getTranslate("CurrenciesHaveBeenImportedUnSuccessfully")}! {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }
    }
}
