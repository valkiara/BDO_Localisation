using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace BDO_Localisation_AddOn
{
    class BDOSFuelConsumption
    {
        public static bool openFormEvent = false;
        public static void createDocumentUDO(out string errorText)
        {
            errorText = null;
            string tableName = "BDOSFUECON";
            string description = "Fuel Consumption Document";
            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_Document, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>(); // DocDate 
            fieldskeysMap.Add("Name", "DocDate");
            fieldskeysMap.Add("TableName", "BDOSFUECON");
            fieldskeysMap.Add("Description", "Posting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FromDate");
            fieldskeysMap.Add("TableName", "BDOSFUECON");
            fieldskeysMap.Add("Description", "From Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ToDate");
            fieldskeysMap.Add("TableName", "BDOSFUECON");
            fieldskeysMap.Add("Description", "To Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);


            //ცხრილური ნაწილი
            tableName = "BDOSFUCON1";
            description = "Fuel Consumption Child1";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "AssetCode");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Asset Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "AssetName");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Asset Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ConsumType");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Consumption Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "StartUnit");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Starting Unit");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Quantity);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "EndUnit");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Ending Unit");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Quantity);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "WorkHours");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Worked Hours");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Quantity);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FuelType");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Fuel Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FuelName");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Fuel Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Uom");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Uom");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "NormConsum");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Normative Consumption");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Consum");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Consumption");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Project");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Project");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension1");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Dimension1");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension2");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Dimension2");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension3");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Dimension3");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension4");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Dimension4");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension5");
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("Description", "Dimension5");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "DocNum");
			fieldskeysMap.Add("TableName", "BDOSFUCON1");
			fieldskeysMap.Add("Description", "DocNum");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 50);

			UDO.addUserTableFields(fieldskeysMap, out errorText);
			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "LineNum");
			fieldskeysMap.Add("TableName", "BDOSFUCON1");
			fieldskeysMap.Add("Description", "LineNum");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 50);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			GC.Collect();

        }
        public static void registerUDO(out string errorText)
        {
            errorText = null;
            string code = "UDO_F_BDOSFUECON_D"; //20 characters (must include at least one alphabetical character).
            Dictionary<string, object> formProperties;

            formProperties = new Dictionary<string, object>();
            formProperties.Add("Name", "Fuel Consumption Document"); //100 characters
            formProperties.Add("TableName", "BDOSFUECON");
            formProperties.Add("ObjectType", SAPbobsCOM.BoUDOObjType.boud_Document);
            formProperties.Add("CanCancel", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanClose", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanCreateDefaultForm", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanDelete", SAPbobsCOM.BoYesNoEnum.tNO);
            formProperties.Add("CanFind", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanYearTransfer", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("ManageSeries", SAPbobsCOM.BoYesNoEnum.tNO);
            formProperties.Add("CanLog", SAPbobsCOM.BoYesNoEnum.tYES);

            List<Dictionary<string, object>> listFindColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listFormColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listChildTables = new List<Dictionary<string, object>>();

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_DocDate");
            fieldskeysMap.Add("ColumnDescription", "Posting Date");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_FromDate");
            fieldskeysMap.Add("ColumnDescription", "From Date");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_ToDate");
            fieldskeysMap.Add("ColumnDescription", "To Date");
            listFindColumns.Add(fieldskeysMap);

            formProperties.Add("FindColumns", listFindColumns);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("FormColumnAlias", "DocEntry");
            fieldskeysMap.Add("FormColumnDescription", "DocEntry");
            listFormColumns.Add(fieldskeysMap);

            formProperties.Add("FormColumns", listFormColumns);

            //ცხრილური ნაწილები
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("TableName", "BDOSFUCON1");
            fieldskeysMap.Add("ObjectName", "BDOSFUCON1");
            listChildTables.Add(fieldskeysMap);

            formProperties.Add("ChildTables", listChildTables);
            //ცხრილური ნაწილები

            UDO.registerUDO(code, formProperties, out errorText);

            GC.Collect();
        }
        public static void addMenus(out string errorText)
        {
            errorText = null;

            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			string query = @"SELECT ""U_BDOSEnbFlM"" AS ""EnableFuelMng"" FROM ""OADM"" WHERE ""U_BDOSEnbFlM"" = 'Y'";

			oRecordSet.DoQuery(query);

			if (!oRecordSet.EoF)
			{
				try
				{
					fatherMenuItem = Program.uiApp.Menus.Item("3072");

					// Add a pop-up menu item
					oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
					oCreationPackage.Checked = false;
					oCreationPackage.Enabled = true;
					oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
					oCreationPackage.UniqueID = "UDO_F_BDOSFUECON_D";
					oCreationPackage.String = BDOSResources.getTranslate("FuelConsumptionDocument");
					oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

					menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
				}
				catch (Exception ex)
				{
					errorText = ex.Message;
				}
			}
        }
        public static void uiApp_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent, out string errorText)
        {
            errorText = null;
            BubbleEvent = true;

            //----------------------------->Cancel<-----------------------------
            try
            {
                SAPbouiCOM.Form oDocForm = Program.uiApp.Forms.ActiveForm;


                if (pVal.BeforeAction == false && pVal.MenuUID == "BDOSAddRow")
                {
                    addMatrixRow(oDocForm, out errorText);
                }
                else if (pVal.BeforeAction == false && pVal.MenuUID == "BDOSDelRow")
                {
                    delMatrixRow(oDocForm, out errorText);
                }
            }
            catch (Exception ex)
            {
                Program.uiApp.MessageBox(ex.ToString(), 1, "", "");
            }
        }
        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == true)
            {
                return;
            }

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "UDO_FT_UDO_F_BDOSFUECON_D")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
                {
                    openFormEvent = false;
                    setVisibleFormItems(oForm, out errorText);
                }
            }

        }
        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
				SAPbouiCOM.DBDataSources oDBDataSources = oForm.DataSources.DBDataSources;

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems(oForm, out BubbleEvent, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                }
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS && oForm.Visible == true && oForm.VisibleEx == true && openFormEvent == false)
                {

                    setVisibleFormItems(oForm, out errorText);

                    string docEntry = oForm.DataSources.DBDataSources.Item("@BDOSFUECON").GetValue("DocEntry", 0).Trim();
                    if (string.IsNullOrEmpty(docEntry))
                    {
                        addMatrixRow(oForm, out errorText);
                    }

                    openFormEvent = true;
                }
				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.ItemUID == "1" && pVal.BeforeAction == true)
				{	
					string DocDates = oDBDataSources.Item("@BDOSFUECON").GetValue("U_DocDate", 0);
					DateTime DocDate = Convert.ToDateTime(DateTime.ParseExact(DocDates, "yyyyMMdd", CultureInfo.InvariantCulture));

					string ToDates = oDBDataSources.Item("@BDOSFUECON").GetValue("U_ToDate", 0);
					DateTime ToDate = Convert.ToDateTime(DateTime.ParseExact(ToDates, "yyyyMMdd", CultureInfo.InvariantCulture));
					SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("FuelMTR").Specific;
					string consumption;
					double cons = 0;
					double endUnit = 0;
					double startUnit = 0;
					string dimension;
					SAPbobsCOM.Recordset oRecordSetDim = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

					string dimQuery = @"SELECT ""U_BDOSFADim"" FROM ""OADM"" ";
					oRecordSetDim.DoQuery(dimQuery);
					string dim = oRecordSetDim.Fields.Item("U_BDOSFADim").Value;
					var dim_Col = "Dimension" + dim;
					
					for (int i = 1; i < oMatrix.RowCount; i++) {
						consumption = oMatrix.Columns.Item("Consum").Cells.Item(i).Specific.Value;
						endUnit = Convert.ToDouble(oMatrix.Columns.Item("EndUnit").Cells.Item(i).Specific.Value);
						startUnit = Convert.ToDouble(oMatrix.Columns.Item("StartUnit").Cells.Item(i).Specific.Value);
						dimension = oMatrix.Columns.Item(dim_Col).Cells.Item(i).Specific.Value;
						if (consumption == "")
						{
							Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("ConsumptionColumnMustBeFilled"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
							BubbleEvent = false;
							break;
						}
						else {
							cons = Convert.ToDouble(consumption);
							if (cons == 0 || cons < 0)
							{
								Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("ConsumptionValueMustBeGreaterThanZero"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
								BubbleEvent = false;
								break;
							}
						}
						if (endUnit == 0 || endUnit < startUnit)
						{
							Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("YouMustEnterEndingUnitGreaterThanStartingUnit"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
							BubbleEvent = false;
							break;
						}
						if (dimension == "") {
							Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("AssetsDimensionMustBefilled"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
							BubbleEvent = false;
							break;
						}
						
					}

					if (DocDate < ToDate) { 
						Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("YouMustEnterPostingDateAfterEndingDate"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
						BubbleEvent = false;
					}
				}

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false)
					{
					oForm.Freeze(true);
					if (Program.FORM_LOAD_FOR_VISIBLE == true)
					{
						setSizeForm(oForm, out errorText);
						oForm.Title = BDOSResources.getTranslate("FuelConsumptionDocument");
						Program.FORM_LOAD_FOR_VISIBLE = false;
					}
					oForm.Freeze(false);
					setVisibleFormItems(oForm, out errorText);
				}

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm(oForm, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));
                    chooseFromList(oForm, oCFLEvento, pVal, out errorText);
                }
			
				if (pVal.ItemUID == "FuelMTR" && pVal.ColUID == "LineID" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK & pVal.BeforeAction == true)
                {
                    BubbleEvent = false;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "Dimension1" || pVal.ItemUID == "Dimension2" || pVal.ItemUID == "Dimension3"
                         || pVal.ItemUID == "Dimension4" || pVal.ItemUID == "Dimension5")
                    {
                        setVisibleFormItems(oForm, out errorText);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false & Program.FORM_LOAD_FOR_VISIBLE == true)
                {
                    oForm.Freeze(true);
                    setSizeForm(oForm, out errorText);
                    oForm.Title = BDOSResources.getTranslate("FuelConsumptionDocument");
                    oForm.Freeze(false);
                    Program.FORM_LOAD_FOR_VISIBLE = false;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm(oForm, out errorText);
                    oForm.Freeze(false);
                }
				if (pVal.ItemUID == "FuelMTR")
				{
					if ((pVal.ColUID == "WorkHours" || pVal.ColUID == "EndUnit") & pVal.ItemChanged & pVal.BeforeAction == false)
					{
						oForm.Freeze(true);
						SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("FuelMTR").Specific;
						double  endUnit = Convert.ToDouble(oMatrix.Columns.Item("EndUnit").Cells.Item(pVal.Row).Specific.Value);
						double workedHours = Convert.ToDouble(oMatrix.Columns.Item("WorkHours").Cells.Item(pVal.Row).Specific.Value);
						//oDBDataSources.Item("@BDOSFUCON1").SetValue("U_EndUnit", pVal.Row-1, endUnit);
						//oDBDataSources.Item("@BDOSFUCON1").SetValue("U_WorkHours", pVal.Row-1,workedHours);
						fillConsumptionAmount(oForm, out errorText, pVal, oMatrix,endUnit,workedHours);
						oForm.Freeze(false);
					}
				}

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE & pVal.BeforeAction == false)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE == true)
                    {
                        oForm.Freeze(true);
                        setVisibleFormItems(oForm, out errorText);
                        openFormEvent = false;
                        oForm.Freeze(false);
                        Program.FORM_LOAD_FOR_ACTIVATE = false;
                    }
                }
            }
        }
        public static void uiApp_RightClickEvent(SAPbouiCOM.Form oForm, SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (eventInfo.ItemUID == "FuelMTR")
            {
                SAPbouiCOM.MenuItem oMenuItem;
                SAPbouiCOM.Menus oMenus;
                SAPbouiCOM.MenuCreationParams oCreationPackage;

                try
                {
                    oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "BDOSAddRow";
                    oCreationPackage.String = BDOSResources.getTranslate("AddNewRow");
                    oCreationPackage.Enabled = true;
                    oCreationPackage.Position = -1;

                    oMenuItem = Program.uiApp.Menus.Item("1280");
                    oMenus = oMenuItem.SubMenus;
                    oMenus.AddEx(oCreationPackage);
                }
                catch (Exception ex)
                {
                    string errorText = ex.Message;
                }

                try
                {
                    oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "BDOSDelRow";
                    oCreationPackage.String = BDOSResources.getTranslate("DeleteRow");
                    oCreationPackage.Enabled = true;
                    oCreationPackage.Position = -1;

                    oMenuItem = Program.uiApp.Menus.Item("1280");
                    oMenus = oMenuItem.SubMenus;
                    oMenus.AddEx(oCreationPackage);
                }
                catch (Exception ex)
                {
                    string errorText = ex.Message;
                }
            }
            else
            {
                try
                {
                    Program.uiApp.Menus.RemoveEx("BDOSAddRow");
                }
                catch { }
                try
                {
                    Program.uiApp.Menus.RemoveEx("BDOSDelRow");
                }
                catch { }
            }
        }
        public static void createFormItems(SAPbouiCOM.Form oForm, out bool BubbleEvent, out string errorText)
        {
            BubbleEvent = true;
            errorText = null;

            Dictionary<string, object> formItems;
            string itemName = "";

            int left_s = 6;
            int left_e = 127;
            int height = 15;
            int top = 6;
            int width_s = 120;
            int width_e = 148;

            top = top + height + 1;
            oForm.AutoManaged = true;

            formItems = new Dictionary<string, object>();
            itemName = "DocDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("PostingDate"));
            formItems.Add("LinkTo", "DocDateE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "DocDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUECON");
            formItems.Add("Alias", "U_DocDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("SetAutoManaged", true);


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "FromDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("FromDate"));
            formItems.Add("LinkTo", "FromDateE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "FromDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUECON");
            formItems.Add("Alias", "U_FromDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("SetAutoManaged", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "ToDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ToDate"));
            formItems.Add("LinkTo", "ToDateE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "ToDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUECON");
            formItems.Add("Alias", "U_ToDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("SetAutoManaged", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "StatusS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Status"));
            formItems.Add("LinkTo", "StatusC");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "StatusC"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUECON");
            formItems.Add("Alias", "Status");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("SetAutoManaged", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "CanceledS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Canceled"));
            formItems.Add("LinkTo", "CanceledC");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "CanceledC"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUECON");
            formItems.Add("Alias", "Canceled");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            left_s = 6;
            left_e = 127;
            top = top + 2 * height + 1;

            string objectTypeItem = "4";
            bool multiSelection = false;

            //ცხრილური ნაწილები
            left_s = 6;
            left_e = 127;
            top = top + 2 * height + 1;

            //მატრიცა
            formItems = new Dictionary<string, object>();
            itemName = "FuelMTR"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            formItems.Add("Left", left_s);
            formItems.Add("Width", oForm.Width);
            formItems.Add("Top", top);
            formItems.Add("Height", 70);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("AffectsFormMode", false);
            formItems.Add("SetAutoManaged", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            string uniqueID_lf_ItemMTR_CFL = "ItemMTR_CFL";
            FormsB1.addChooseFromList(oForm, true, objectTypeItem, uniqueID_lf_ItemMTR_CFL);

            //პირობის დადება ძს არჩევის სიაზე
            SAPbouiCOM.ChooseFromList oCFL_Item = oForm.ChooseFromLists.Item(uniqueID_lf_ItemMTR_CFL);
            SAPbouiCOM.Conditions oCons_Item = oCFL_Item.GetConditions();
            SAPbouiCOM.Condition oCon_Item = oCons_Item.Add();
            oCon_Item.Alias = "ItemType";
            oCon_Item.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon_Item.CondVal = "F"; //Fixed Assets
            oCFL_Item.SetConditions(oCons_Item);

            string objectType = "63";
            string uniqueID_lf_Project = "Project_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_Project);

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("FuelMTR").Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;

            oColumn = oColumns.Add("LineID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "  ";
            oColumn.Width = 20;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("AssetCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("AssetCode");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;
            oColumn.ExtendedObject.LinkedObjectType = objectTypeItem;

            oColumn = oColumns.Add("AssetName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("AssetName");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("ConsumType", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ConsumType");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("StartUnit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("StartingUnit");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("EndUnit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("EndingUnit");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("WorkHours", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("WorkedHours");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("FuelType", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("FuelType");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;
			oColumn.ExtendedObject.LinkedObjectType = objectTypeItem;

			oColumn = oColumns.Add("FuelName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("FuelName");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("Uom", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomEntry");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("NormConsum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("NormativeConsumption");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("Consum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Consum");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("Project", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Project");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;
            oColumn.ExtendedObject.LinkedObjectType = objectType;

            oColumn = oColumns.Add("Dimension1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Dimension1");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("Dimension2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Dimension2");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("Dimension3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Dimension3");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("Dimension4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Dimension4");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("Dimension5", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Dimension5");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

			oColumn = oColumns.Add("DocNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocNum");
			oColumn.Width = 60;
			oColumn.Editable = false;
			oColumn.Visible = false;

			oColumn = oColumns.Add("LineNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			oColumn.TitleObject.Caption = BDOSResources.getTranslate("LineNum");
			oColumn.Width = 60;
			oColumn.Editable = false;
			oColumn.Visible = false;

			SAPbouiCOM.DBDataSource oDBDataSource;
            oDBDataSource = oForm.DataSources.DBDataSources.Add("@BDOSFUCON1");

            oColumn = oColumns.Item("LineID");
            oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "LineID");

            oColumn = oColumns.Item("AssetCode");
            oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_AssetCode");
            oColumn.ChooseFromListUID = uniqueID_lf_ItemMTR_CFL;
            oColumn.ChooseFromListAlias = "ItemCode";

            oColumn = oColumns.Item("AssetName");
            oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_AssetName");


            oColumn = oColumns.Item("ConsumType");
            oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_ConsumType");

            oColumn = oColumns.Item("StartUnit");
            oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_StartUnit");

            oColumn = oColumns.Item("EndUnit");
            oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_EndUnit");

            oColumn = oColumns.Item("WorkHours");
            oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_WorkHours");

            oColumn = oColumns.Item("FuelType");
            oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_FuelType");

            oColumn = oColumns.Item("FuelName");
            oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_FuelName");

            oColumn = oColumns.Item("Uom");
            oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_Uom");

            oColumn = oColumns.Item("NormConsum");
            oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_NormConsum");

            oColumn = oColumns.Item("Consum");
            oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_Consum");

            oColumn = oColumns.Item("Project");
            oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_Project");
            oColumn.ChooseFromListUID = uniqueID_lf_Project;
            oColumn.ChooseFromListAlias = "PrjCode";
            for (int i = 1; i <= 5; i++)
            {
                objectType = "62";
                string uniqueID_lf_DistrRule = "Rule_CFL" + i.ToString() + "A";
                FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_DistrRule);

                oColumn = oColumns.Item("Dimension" + i);
                oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_Dimension" + i);
                oColumn.ChooseFromListUID = uniqueID_lf_DistrRule;
                oColumn.ChooseFromListAlias = "OcrCode";
            }
			oColumn = oColumns.Item("DocNum");
			oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_DocNum");

			oColumn = oColumns.Item("LineNum");
			oColumn.DataBind.SetBound(true, "@BDOSFUCON1", "U_LineNum");
                        top = top + 5 + 70;

            //შემქმნელი
            formItems = new Dictionary<string, object>();
            itemName = "CreatorS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Creator"));
            formItems.Add("LinkTo", "CreatorE");


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "CreatorE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUECON");
            formItems.Add("Alias", "Creator");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("SetAutoManaged", true);


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //შენიშვნა
            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "RemarksS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Remarks"));
            formItems.Add("LinkTo", "RemarksE");


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "RemarksE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUECON");
            formItems.Add("Alias", "Remark");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", 3 * height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ScrollBars", SAPbouiCOM.BoScrollBars.sb_Vertical);


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }
        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);

            try
            {
                oForm.Items.Item("CanceledC").Enabled = false;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }
        public static void resizeForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                reArrangeFormItems(oForm);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }
        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {

            SAPbouiCOM.Item oItem = null;
            int height = 15;
            int top = 6;

            top = top + height + 1;
            oForm.Items.Item("DocDateS").Top = top;
            oForm.Items.Item("DocDateE").Top = top;

            top = top + height + 1;
            oForm.Items.Item("FromDateS").Top = top;
            oForm.Items.Item("FromDateE").Top = top;

            top = top + height + 1;
            oForm.Items.Item("ToDateS").Top = top;
            oForm.Items.Item("ToDateE").Top = top;

            top = top + height + 1;
            oForm.Items.Item("StatusS").Top = top;
            oForm.Items.Item("StatusC").Top = top;

            top = top + height + 1;
            oForm.Items.Item("CanceledS").Top = top;
            oForm.Items.Item("CanceledC").Top = top;



            top = top + 2 * height + 1;


            int MTRWidth = oForm.Width - 15;
            top = top + height + 1;
            oItem = oForm.Items.Item("FuelMTR");
            oItem.Top = top;
            oItem.Width = MTRWidth;
            oItem.Height = oForm.Height / 3;

            // სვეტების ზომები 
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("FuelMTR").Specific));
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;
            oColumn = oMatrix.Columns.Item("LineID");
            oColumn.Width = 20 - 1;
            MTRWidth = MTRWidth - 20 - 1;

            //სარდაფი
            top = oItem.Top + oItem.Height + 20;

            oItem = oForm.Items.Item("CreatorS");
            oItem.Top = top;
            oItem = oForm.Items.Item("CreatorE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("RemarksS");
            oItem.Top = top;
            oItem = oForm.Items.Item("RemarksE");
            oItem.Top = top;

            top = top + 4 * height;

            oForm.Items.Item("1").Top = top;
            oForm.Items.Item("2").Top = top;


        }
		public static void setSizeForm(SAPbouiCOM.Form oForm, out string errorText)
		{
			errorText = null;
			try
			{
				oForm.ClientHeight = Program.uiApp.Desktop.Width / 3;
				oForm.Height = Program.uiApp.Desktop.Width / 2;

				oForm.Left = (Program.uiApp.Desktop.Width - oForm.Width) / 2;
				oForm.Top = (Program.uiApp.Desktop.Height - oForm.Height) / 2;
			}
			catch (Exception ex)
			{
				errorText = ex.Message;
			}
		}
		public static void fillConsumptionAmount(SAPbouiCOM.Form oForm, out string errorText, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.Matrix oMatrix, double endUnit, double workedHours)
		{
			errorText = null;

			oMatrix = oForm.Items.Item("FuelMTR").Specific;
			double perH;
			double perK;
			string assetCode = oMatrix.Columns.Item("AssetCode").Cells.Item(pVal.Row).Specific.Value;
			double startUnit = Convert.ToDouble(oMatrix.Columns.Item("StartUnit").Cells.Item(pVal.Row).Specific.Value);

			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			string query = @"SELECT ""U_PerHr"", ""U_PerKm"" FROM ""OITM"" WHERE ""ItemCode"" = '" + assetCode + @"'";
			oRecordSet.DoQuery(query);

			if (!oRecordSet.EoF)
			{
				perH = oRecordSet.Fields.Item("U_PerHr").Value;
				perK = oRecordSet.Fields.Item("U_PerKm").Value;
			}

			else
			{
				perH = 0;
				perK = 0;
			}

			if (endUnit > 0 || workedHours > 0)
			{
				double normCon = perK * ((endUnit - startUnit) / 100) + perH * workedHours;

				try
				{

					SAPbouiCOM.EditText normConsum = oMatrix.Columns.Item("NormConsum").Cells.Item(pVal.Row).Specific;
					normConsum.Value = Convert.ToString(normCon);

					SAPbouiCOM.EditText consum = oMatrix.Columns.Item("Consum").Cells.Item(pVal.Row).Specific;
					consum.Value = Convert.ToString(normCon);

				}

				catch (Exception ex)
				{
					errorText = ex.Message;
				}

			}
		}
		public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            bool beforeAction = pVal.BeforeAction;
            int row = pVal.Row;
            string sCFL_ID = oCFLEvento.ChooseFromListUID;

            try
            {
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                SAPbouiCOM.DBDataSources oDBDataSources = oForm.DataSources.DBDataSources;

                if (beforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "ItemMTR_CFL")
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("FuelMTR").Specific));
                            oMatrix.FlushToDataSource();

                            oDBDataSources.Item("@BDOSFUCON1").SetValue("U_AssetCode", pVal.Row - 1, oDataTable.GetValue("ItemCode", 0));
                            oDBDataSources.Item("@BDOSFUCON1").SetValue("U_AssetName", pVal.Row - 1, oDataTable.GetValue("ItemName", 0));
                            oDBDataSources.Item("@BDOSFUCON1").SetValue("U_ConsumType", pVal.Row - 1, oDataTable.GetValue("U_FltCode", 0));
                            oDBDataSources.Item("@BDOSFUCON1").SetValue("U_FuelType", pVal.Row - 1, oDataTable.GetValue("U_FuelType", 0));
                            oDBDataSources.Item("@BDOSFUCON1").SetValue("U_FuelName", pVal.Row - 1, oDataTable.GetValue("U_Name", 0));
                            oDBDataSources.Item("@BDOSFUCON1").SetValue("U_Uom", pVal.Row - 1, oDataTable.GetValue("U_UomCode", 0));

                            double unit;
							string assCode = oDataTable.GetValue("ItemCode", 0);

							SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            string query = @"SELECT TOP 1 ""U_EndUnit"" FROM ""@BDOSFUCON1"" WHERE ""U_AssetCode"" = '" + assCode + @"' ORDER BY ""U_EndUnit"" DESC";
                            oRecordSet.DoQuery(query);


                            if (!oRecordSet.EoF)
                            {
                                unit = oRecordSet.Fields.Item("U_EndUnit").Value;
                            }

                            else
                            {
                                unit = 0;
                            }

                            oDBDataSources.Item("@BDOSFUCON1").SetValue("U_StartUnit", pVal.Row - 1, unit.ToString());


                            SAPbobsCOM.Recordset oRecordSetDim = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            string dimQuery = @"SELECT ""U_BDOSFADim"" FROM ""OADM"" ";
                            oRecordSetDim.DoQuery(dimQuery);
                            string dim = oRecordSetDim.Fields.Item("U_BDOSFADim").Value;

                            string code = oDataTable.GetValue("ItemCode", 0);

                            SAPbobsCOM.Recordset oRecordSet2 = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            string query2 = 
                                @"SELECT
                                      ""@BDOSFUCON1"".""U_AssetCode"" as ""assetcode"",
                                      ""OPRC"".""U_BDOSFACode"" as ""code"",
                                      ""OPRC"".""PrcCode"" as ""pCode""
                                FROM 
                                      ""OPRC""
                                LEFT JOIN 
                                      ""@BDOSFUCON1"" 
                                ON 
                                      ""OPRC"".""U_BDOSFACode"" = ""@BDOSFUCON1"".""U_AssetCode"" 
                                WHERE ""OPRC"".""U_BDOSFACode""  = '" + code + @"'";

                            oRecordSet2.DoQuery(query2);
                            if (!oRecordSet2.EoF)
                            {
                                var dim_Col = "U_Dimension" + dim; 
                                oDBDataSources.Item("@BDOSFUCON1").SetValue(dim_Col, pVal.Row - 1, oRecordSet2.Fields.Item("pCode").Value);
                            }
                            
                            oMatrix.LoadFromDataSource();
                            addMatrixRow(oForm, out errorText);
                        }

                        else if (sCFL_ID == "Project_CFL")
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("FuelMTR").Specific));
                            oMatrix.FlushToDataSource();
                            oDBDataSources.Item("@BDOSFUCON1").SetValue("U_Project", pVal.Row - 1, oDataTable.GetValue("PrjCode", 0));
                            oMatrix.LoadFromDataSource();
                        }

                        else if (sCFL_ID.Length >= 2 && sCFL_ID.Substring(0, sCFL_ID.Length - 2) == "Rule_CFL")
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("FuelMTR").Specific));
                            oMatrix.FlushToDataSource();
                            SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                            string val = oDataTableSelectedObjects.GetValue("OcrCode", 0);
                            oDBDataSources.Item("@BDOSFUCON1").SetValue("U_Dimension" + sCFL_ID.Substring(sCFL_ID.Length - 2, 1), pVal.Row - 1, oDataTable.GetValue("OcrCode", 0));
                            oMatrix.LoadFromDataSource();
                        }

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                }
                else
                {
                    if (sCFL_ID.Length > 1 && sCFL_ID.Substring(0, sCFL_ID.Length - 2) == "Rule_CFL")
                    {
                        oForm.Freeze(true);
                        string dimensionCode = sCFL_ID.Substring(sCFL_ID.Length - 2, 1);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        string strDocDate = oDBDataSources.Item("@BDOSFUECON").GetValue("U_DocDate", 0);
                        DateTime DocDate = DateTime.TryParseExact(strDocDate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DocDate) == false ? DateTime.Now : DocDate;

                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = @"SELECT
	                                     ""OCR1"".""OcrCode"",
	                                     ""OOCR"".""DimCode"" 
                                    FROM ""OCR1"" 
                                    LEFT JOIN ""OOCR"" ON ""OCR1"".""OcrCode"" = ""OOCR"".""OcrCode"" 
                                    WHERE ""OOCR"".""DimCode"" = " + dimensionCode + @" AND ""ValidFrom"" <= '" + DocDate.ToString("yyyyMMdd") +
                                                                                         @"' AND (""ValidTo"" > '" + DocDate.ToString("yyyyMMdd") + @"' OR " + @" ""ValidTo"" IS NULL)";

                        try
                        {
                            oRecordSet.DoQuery(query);
                            int recordCount = oRecordSet.RecordCount;
                            int i = 1;

                            while (!oRecordSet.EoF)
                            {
                                SAPbouiCOM.Condition oCon = oCons.Add();
                                oCon.Alias = "OcrCode";
                                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCon.CondVal = oRecordSet.Fields.Item("OcrCode").Value.ToString();
                                oCon.Relationship = (i == recordCount) ? SAPbouiCOM.BoConditionRelationship.cr_NONE : SAPbouiCOM.BoConditionRelationship.cr_OR;

                                i = i + 1;
                                oRecordSet.MoveNext();
                            }

                            //თუ არცერთი შეესაბამება ცარიელზე გავიდეს
                            if (oCons.Count == 0)
                            {
                                SAPbouiCOM.Condition oCon = oCons.Add();
                                oCon.Alias = "OcrCode";
                                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCon.CondVal = "";
                            }

                            oCFL.SetConditions(oCons);
                        }
                        catch (Exception ex)
                        {
                            errorText = ex.Message;
                        }

                        oForm.Freeze(false);
                    }

                    else if (sCFL_ID == "ItemMTR_CFL")

                    {
                        oForm.Freeze(true);
                        try
                        {
                            string queryVis = @"SELECT ""ItemCode""
                                                From ""OITM""
                                                INNER JOIN ""OACS""
                                                ON ""OITM"".""AssetClass"" = ""OACS"".""Code"" AND ""OACS"".""U_visCode"" = 'Y'";

                            SAPbobsCOM.Recordset oRecordSetVis = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            oRecordSetVis.DoQuery(queryVis);
                            SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                            int i = 1;
                            int recordCount = oRecordSetVis.RecordCount;
                            while (!oRecordSetVis.EoF)
                            {
                                SAPbouiCOM.Condition oCon = oCons.Add();
                                oCon.Alias = "ItemCode";
                                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCon.CondVal = oRecordSetVis.Fields.Item("ItemCode").Value.ToString();
                                oCFL.SetConditions(oCons);
                                oCon.Relationship = (i == recordCount) ? SAPbouiCOM.BoConditionRelationship.cr_NONE : SAPbouiCOM.BoConditionRelationship.cr_OR;
                                i = i + 1;
                                oRecordSetVis.MoveNext();
                            }
                            oCFL.SetConditions(oCons);
                        }
                        catch (Exception ex)
                        {
                            errorText = ex.Message;
                        }

                        oForm.Freeze(false);
                    }
                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
            }
            finally
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                GC.Collect();
            }
        }
        public static void addMatrixRow(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("FuelMTR").Specific));
            SAPbouiCOM.DBDataSource mtrDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFUCON1");

            oMatrix.FlushToDataSource();
            if (mtrDataSource.GetValue("U_AssetCode", mtrDataSource.Size - 1) != "")
            {
                mtrDataSource.InsertRecord(mtrDataSource.Size);
            }
            mtrDataSource.SetValue("LineId", mtrDataSource.Size - 1, mtrDataSource.Size.ToString());

            oMatrix.LoadFromDataSource();

            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }

            oForm.Freeze(false);
        }
        public static void delMatrixRow(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            try
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("FuelMTR").Specific));
                SAPbouiCOM.DBDataSource mtrDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFUCON1");

                oMatrix.FlushToDataSource();
                int firstRow = 0;
                int row = 0;
                int deletedRowCount = 0;

                while (row != -1)
                {
                    row = oMatrix.GetNextSelectedRow(firstRow, SAPbouiCOM.BoOrderType.ot_RowOrder);
                    if (row > -1)
                    {
                        deletedRowCount++;
                        mtrDataSource.RemoveRecord(row - deletedRowCount);
                        firstRow = row;
                    }
                }

                for (int i = 0; i <= mtrDataSource.Size; i++)
                {
                    mtrDataSource.SetValue("LineId", i, (i + 1).ToString());
                }

                oMatrix.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && deletedRowCount > 0)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
                oForm.Freeze(false);
            }

        }
    }
}


