using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;

namespace BDO_Localisation_AddOn
{
    static partial class BDO_ImportRateForm
    {
        public static void createForm(SAPbouiCOM.Form oExchangeFormRatesAndIndexes, out string errorText)
        {
            errorText = null;

            int yearNumber = Convert.ToInt32(oExchangeFormRatesAndIndexes.DataSources.UserDataSources.Item(1).ValueEx);
            int monthNumber = Convert.ToInt32(oExchangeFormRatesAndIndexes.DataSources.UserDataSources.Item(4).ValueEx);

            DateTime startOfMonth = new DateTime(yearNumber, monthNumber, 1);
            DateTime endOfMonth = new DateTime(yearNumber, monthNumber, DateTime.DaysInMonth(yearNumber, monthNumber));
            endOfMonth = endOfMonth > DateTime.Today ? DateTime.Today : endOfMonth;

            if (startOfMonth > DateTime.Today)
            {
                Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("EnterStartEndDates") + " " + startOfMonth + "-" + endOfMonth, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return;
            }

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDO_ImportRateForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("ImportRate"));
            formProperties.Add("Left", 558);
            formProperties.Add("ClientWidth", 400);
            formProperties.Add("Top", 335);
            formProperties.Add("ClientHeight", 299);
            formProperties.Add("Modality", SAPbouiCOM.BoFormModality.fm_Modal);

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
                    //ფორმის ელემენტების თვისებები
                    Dictionary<string, object> formItems = null;

                    string itemName = "";
                    int left = 6;
                    formItems = new Dictionary<string, object>();
                    itemName = "dateFrom";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 50);
                    formItems.Add("Top", 30);
                    formItems.Add("Caption", BDOSResources.getTranslate("StartDate"));
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "startDate");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 50 + 10;

                    string startOfMonthStr = startOfMonth.ToString("yyyyMMdd");
                    formItems = new Dictionary<string, object>();
                    itemName = "startDate";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", 30);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", startOfMonthStr);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 10;

                    itemName = "dateTo";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 50);
                    formItems.Add("Top", 30);
                    formItems.Add("Caption", "-");
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "endDate");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 50 + 10;

                    string endOfMonthStr = endOfMonth.ToString("yyyyMMdd");
                    formItems = new Dictionary<string, object>();
                    itemName = "endDate";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", 30);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", endOfMonthStr);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    itemName = "3";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", 5);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", 260);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", "OK");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    itemName = "2";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", 75);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", 260);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Cancel"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    itemName = "currMatrix";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", 5);
                    formItems.Add("Width", 400);
                    formItems.Add("Top", 50);
                    formItems.Add("Height", 200);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("currMatrix").Specific));
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;

                    oColumn = oColumns.Add("DSYN", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    oColumn.TitleObject.Caption = "";
                    oColumn.Width = 30;
                    oColumn.Editable = true;
                    oColumn.ValOff = "N";
                    oColumn.ValOn = "Y";

                    oColumn = oColumns.Add("DSCurrName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Currency");
                    oColumn.Width = 40;
                    oColumn.Editable = false;

                    oColumn = oColumns.Add("DSCurrCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("CurrencyCode");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.Visible = false;

                    SAPbouiCOM.UserDataSource oUserDataSource;
                    SAPbouiCOM.DBDataSource oDBDataSource;
                    oUserDataSource = oForm.DataSources.UserDataSources.Add("YN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                    oDBDataSource = oForm.DataSources.DBDataSources.Add("OCRN");

                    oColumn = oColumns.Item("DSCurrName");
                    oColumn.DataBind.SetBound(true, "OCRN", "CurrName");
                    oColumn = oColumns.Item("DSYN");
                    oColumn.DataBind.SetBound(true, "", "YN");
                    oColumn = oColumns.Item("DSCurrCode");
                    oColumn.DataBind.SetBound(true, "OCRN", "CurrCode");

                    string MainCurncy = CurrencyB1.getMainCurrency(out errorText);

                    oMatrix.Clear();
                    SAPbouiCOM.Conditions oConds = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition oCond = oConds.Add();
                    oCond.Alias = "CurrCode";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                    oCond.CondVal = MainCurncy;
                    oDBDataSource.Query(oConds);
                    oUserDataSource.ValueEx = "N";
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();
                }

                oForm.Visible = true;
            }

            GC.Collect();
        }

        public static bool date_OnChanged(SAPbouiCOM.Form oForm, SAPbouiCOM.Form oExchangeFormRatesAndIndexes, out string errorText)
        {
            errorText = null;

            try
            {
                int yearNumber = Convert.ToInt32(oExchangeFormRatesAndIndexes.DataSources.UserDataSources.Item(1).ValueEx);
                int monthNumber = Convert.ToInt32(oExchangeFormRatesAndIndexes.DataSources.UserDataSources.Item(4).ValueEx);

                DateTime startOfMonth = new DateTime(yearNumber, monthNumber, 1);
                DateTime endOfMonth = new DateTime(yearNumber, monthNumber, DateTime.DaysInMonth(yearNumber, monthNumber));
                endOfMonth = endOfMonth > DateTime.Today ? DateTime.Today : endOfMonth;

                string startDateStr = oForm.DataSources.UserDataSources.Item("startDate").ValueEx;
                DateTime startDate = FormsB1.DateFormats(startDateStr, "yyyyMMdd") == new DateTime() ? DateTime.Today : FormsB1.DateFormats(startDateStr, "yyyyMMdd");

                string endDateStr = oForm.DataSources.UserDataSources.Item("endDate").ValueEx;
                DateTime endDate = FormsB1.DateFormats(endDateStr, "yyyyMMdd") == new DateTime() ? DateTime.Today : FormsB1.DateFormats(endDateStr, "yyyyMMdd");

                if (startDate > endDate)
                {
                    Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("EndDate") + " " + BDOSResources.getTranslate("CannotBeEarlierThan") + " " + BDOSResources.getTranslate("StartDate"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return false;
                }

                if (startDate < startOfMonth || endDate > endOfMonth)
                {
                    Program.uiApp.MessageBox(BDOSResources.getTranslate("StartDate") + ": " + startDate + ", " + BDOSResources.getTranslate("EndDate") + ": " + endDate);
                    Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("EnterStartEndDates") + " " + startOfMonth + " - " + endOfMonth, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return false;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return false;
            }
            finally
            {
                GC.Collect();
            }
            return true;
        }

        public static void oK_OnClick(SAPbouiCOM.Form oForm, SAPbouiCOM.Form oExchangeFormRatesAndIndexes)
        {
            if (date_OnChanged(oForm, oExchangeFormRatesAndIndexes, out var errorText) == false)
                return;

            string sXML = null;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("currMatrix").Specific;
            sXML = oMatrix.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(string.Format("<root>{0}</root>", sXML));

            List<string> currencyList = new List<string>();

            foreach (XmlElement node in xmlDoc.GetElementsByTagName("Row"))
            {
                string isImported = node.GetElementsByTagName("Column").Item(0).ChildNodes.Item(1).InnerText;
                string currCode = node.GetElementsByTagName("Column").Item(2).ChildNodes.Item(1).InnerText;
                if (isImported == "Y")
                    currencyList.Add(currCode);
            }

            string startDateStr = oForm.DataSources.UserDataSources.Item("startDate").ValueEx;
            string endDateStr = oForm.DataSources.UserDataSources.Item("endDate").ValueEx;

            oForm.Close();
            Marshal.ReleaseComObject(oForm);

            if (currencyList.Count != 0)
            {
                try
                {
                    //DateTime startDate = FormsB1.DateFormats(startDateStr, "yyyyMMdd") == new DateTime() ? DateTime.Today : FormsB1.DateFormats(startDateStr, "yyyyMMdd");
                    //DateTime endDate = FormsB1.DateFormats(endDateStr, "yyyyMMdd") == new DateTime() ? DateTime.Today : FormsB1.DateFormats(endDateStr, "yyyyMMdd");

                    Dictionary<string, Dictionary<int, double>> currencyListFromNBG = new Dictionary<string, Dictionary<int, double>>();

                    //while (startDate <= endDate)
                    //{
                    startDateStr = DateTime.Today.ToString("yyyy-MM-dd"); //startDate.ToString("yyyy-MM-dd");
                    CurrencyB1.importCurrencyRate(startDateStr, ref currencyListFromNBG, currencyList);
                    //startDate = startDate.AddDays(1);
                    //}

                    oExchangeFormRatesAndIndexes.Freeze(true);

                    SAPbouiCOM.Matrix oMatrix1 = (SAPbouiCOM.Matrix)oExchangeFormRatesAndIndexes.Items.Item("4").Specific;
                    SAPbouiCOM.EditText oEdit;

                    foreach (SAPbouiCOM.Column column in oMatrix1.Columns)
                    {
                        if (column.Editable == true & column.Visible)
                        {
                            string title = column.Title;
                            title = CommonFunctions.getCurrencyInternationalCode(title);

                            if (currencyListFromNBG.Keys.Contains(title))
                            {
                                foreach (KeyValuePair<int, double> dailyRate in currencyListFromNBG[title])
                                {
                                    int key = dailyRate.Key;
                                    double rate = dailyRate.Value;
                                    NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
                                    oEdit = column.Cells.Item(key).Specific;
                                    oEdit.Value = rate.ToString(Nfi);
                                }
                            }
                        }
                    }
                    Program.uiApp.SetStatusBarMessage($"{BDOSResources.getTranslate("CurrenciesHaveBeenImportedSuccessfully")}!", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                catch (Exception ex)
                {
                    Program.uiApp.SetStatusBarMessage($"{BDOSResources.getTranslate("CurrenciesHaveBeenImportedUnSuccessfully")}! {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                finally
                {
                    oExchangeFormRatesAndIndexes.Freeze(false);
                }
                //sXML = oMatrix1.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);
                //XmlDocument xmlDoc2 = new XmlDocument();
                //xmlDoc2.LoadXml(string.Format("<root>{0}</root>", sXML));

                //foreach (XmlElement node in xmlDoc2.GetElementsByTagName("ColumnInfo"))
                //{
                //    string title = node.GetElementsByTagName("Title").Item(0).InnerText;
                //    if (currencyListFromNBG.Keys.Contains(title) == true)
                //    {
                //        string uniqueID = node.GetElementsByTagName("UniqueID").Item(0).InnerText;
                //        foreach (KeyValuePair<int, double> dailyRate in currencyListFromNBG[title])
                //        {
                //            int key = dailyRate.Key;
                //            double rate = dailyRate.Value;

                //            foreach (XmlElement nodeRow in xmlDoc2.GetElementsByTagName("Row")[key - 1])
                //            {
                //                foreach (XmlElement nodeColumn in nodeRow.GetElementsByTagName("Column"))
                //                {
                //                    if (nodeColumn.GetElementsByTagName("ID").Item(0).InnerText == uniqueID)
                //                    {
                //                        NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
                //                        nodeColumn.GetElementsByTagName("Value").Item(0).InnerText = rate.ToString(Nfi);
                //                        oEdit = oMatrix1.Columns.Item(uniqueID).Cells.Item(key).Specific;
                //                        oEdit.Value = rate.ToString(Nfi);
                //                    }
                //                }
                //            }
                //        }
                //    }
                //}
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                int yearNumber = Convert.ToInt32(Program.oExchangeFormRatesAndIndexes.DataSources.UserDataSources.Item(1).ValueEx);
                int monthNumber = Convert.ToInt32(Program.oExchangeFormRatesAndIndexes.DataSources.UserDataSources.Item(4).ValueEx);
                DateTime startOfMonth = new DateTime(yearNumber, monthNumber, 1);
                DateTime endOfMonth = new DateTime(yearNumber, monthNumber, DateTime.DaysInMonth(yearNumber, monthNumber));
                endOfMonth = endOfMonth > DateTime.Today ? DateTime.Today : endOfMonth;

                if (pVal.ItemUID == "startDate" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE & pVal.BeforeAction == false & pVal.InnerEvent == false)
                {
                    if (BDO_ImportRateForm.date_OnChanged(oForm, Program.oExchangeFormRatesAndIndexes, out errorText) == false)
                    {
                        SAPbouiCOM.EditText oStartDate = ((SAPbouiCOM.EditText)(oForm.Items.Item("startDate").Specific));
                        oStartDate.Value = startOfMonth.ToString("yyyyMMdd");
                        return;
                    }
                }

                if (pVal.ItemUID == "endDate" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE & pVal.BeforeAction == false & pVal.InnerEvent == false)
                {
                    if (date_OnChanged(oForm, Program.oExchangeFormRatesAndIndexes, out errorText) == false)
                    {
                        SAPbouiCOM.EditText oEndDate = ((SAPbouiCOM.EditText)(oForm.Items.Item("endDate").Specific));
                        oEndDate.Value = endOfMonth.ToString("yyyyMMdd");
                        return;
                    }
                }

                if (pVal.ItemUID == "3" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    oK_OnClick(oForm, Program.oExchangeFormRatesAndIndexes);
                    Program.oExchangeFormRatesAndIndexes = null;
                }
            }
        }
    }
}
