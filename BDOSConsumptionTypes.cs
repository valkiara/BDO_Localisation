using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
	class BDOSConsumptionTypes
	{
		public static void createMasterDataUDO(out string errorText)
		{
			errorText = null;
			string tableName = "BDOSFLTP";
			string description = "Fleet Types";

			SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
			oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
			System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

			int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_MasterData, out errorText);

			if (result != 0)
			{
				return;
			}

			Dictionary<string, object> fieldskeysMap;
			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "FuelType");
			fieldskeysMap.Add("TableName", "BDOSFLTP");
			fieldskeysMap.Add("Description", "Fuel Type");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 50);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "Name");
			fieldskeysMap.Add("TableName", "BDOSFLTP");
			fieldskeysMap.Add("Description", "Name");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 100);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "UomCode");
			fieldskeysMap.Add("TableName", "BDOSFLTP");
			fieldskeysMap.Add("Description", "Uom Code");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 20);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "PerKm");
			fieldskeysMap.Add("TableName", "BDOSFLTP");
			fieldskeysMap.Add("Description", "Per 100 km");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
			fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Quantity);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "PerHr");
			fieldskeysMap.Add("TableName", "BDOSFLTP");
			fieldskeysMap.Add("Description", "Per hour");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
			fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Quantity);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			GC.Collect();
		}
		public static void registerUDO(out string errorText)
		{
			errorText = null;
			string code = "UDO_F_BDOSFLTP_T"; //20 characters (must include at least one alphabetical character).
			Dictionary<string, object> formProperties;

			formProperties = new Dictionary<string, object>();
			formProperties.Add("Name", "Fleet Types"); //100 characters
			formProperties.Add("TableName", "BDOSFLTP");
			formProperties.Add("ObjectType", SAPbobsCOM.BoUDOObjType.boud_MasterData);
			formProperties.Add("CanCancel", SAPbobsCOM.BoYesNoEnum.tYES);
			formProperties.Add("CanClose", SAPbobsCOM.BoYesNoEnum.tYES);
			formProperties.Add("CanCreateDefaultForm", SAPbobsCOM.BoYesNoEnum.tYES);
			formProperties.Add("CanDelete", SAPbobsCOM.BoYesNoEnum.tNO);
			formProperties.Add("CanFind", SAPbobsCOM.BoYesNoEnum.tYES);
			formProperties.Add("CanYearTransfer", SAPbobsCOM.BoYesNoEnum.tYES);

			List<Dictionary<string, object>> listFindColumns = new List<Dictionary<string, object>>();
			List<Dictionary<string, object>> listFormColumns = new List<Dictionary<string, object>>();

			Dictionary<string, object> fieldskeysMap;

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("ColumnAlias", "Code");
			fieldskeysMap.Add("ColumnDescription", "Code"); //30 characters
			listFindColumns.Add(fieldskeysMap);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("ColumnAlias", "Name");
			fieldskeysMap.Add("ColumnDescription", "Name"); //30 characters
			listFindColumns.Add(fieldskeysMap);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("ColumnAlias", "U_FuelType");
			fieldskeysMap.Add("ColumnDescription", "Fuel Type"); //30 characters
			listFindColumns.Add(fieldskeysMap);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("ColumnAlias", "U_UomCode");
			fieldskeysMap.Add("ColumnDescription", "Uom Code"); //30 characters
			listFindColumns.Add(fieldskeysMap);

			formProperties.Add("FindColumns", listFindColumns);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
			fieldskeysMap.Add("FormColumnAlias", "Code");
			fieldskeysMap.Add("FormColumnDescription", "Code"); //30 characters
			listFormColumns.Add(fieldskeysMap);

			formProperties.Add("FormColumns", listFormColumns);

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
					oCreationPackage.UniqueID = "UDO_F_BDOSFLTP_T";
					oCreationPackage.String = BDOSResources.getTranslate("ConsumTypesMasterData");
					oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

					menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
				}
				catch (Exception ex)
				{
					errorText = ex.Message;
				}
		}

		}
		public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
		{
			errorText = null;
			Dictionary<string, object> formItems;

			string itemName = "";

			int left_s = 6;
			int left_e = 127;
			int height = 15;
			int top = 6;
			int width_s = 121;
			int width_e = 148;

			top = top + height + 1;

			bool multiSelection = false;
			string objectType = "4";
			string uniqueIDItemsList_CFL = "StockItems_CFL";
			FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueIDItemsList_CFL);

			//პირობის დადება აითემის არჩევის სიაზე

			SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueIDItemsList_CFL);
			SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
			SAPbouiCOM.Condition oCon = oCons.Add();
			oCon.Alias = "InvntItem";
			oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
			oCon.CondVal = "Y";
			oCFL.SetConditions(oCons);

			formItems = new Dictionary<string, object>();
			itemName = "FuelTypeS"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
			formItems.Add("Left", left_s);
			formItems.Add("Width", width_s);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("Caption", BDOSResources.getTranslate("FuelType"));
			formItems.Add("LinkTo", "FuelTypeE");

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "FuelTypeE"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "@BDOSFLTP");
			formItems.Add("Alias", "U_FuelType");
			formItems.Add("Bound", true);
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			formItems.Add("Left", left_e);
			formItems.Add("Width", width_e);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("DisplayDesc", true);
			formItems.Add("ChooseFromListUID", uniqueIDItemsList_CFL);
			formItems.Add("ChooseFromListAlias", "ItemCode");

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "FuelTyLB"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
			formItems.Add("Left", left_e - 20);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("LinkTo", "FuelTypeE");
			formItems.Add("LinkedObjectType", objectType);
			formItems.Add("FromPane", 0);
			formItems.Add("ToPane", 0);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			top = top + height + 1;

			formItems = new Dictionary<string, object>();
			itemName = "ItemNameS"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
			formItems.Add("Left", left_s);
			formItems.Add("Width", width_s);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("Caption", BDOSResources.getTranslate("Name"));
			formItems.Add("LinkTo", "ItemNameE");

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "ItemNameE"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "@BDOSFLTP");
			formItems.Add("Alias", "U_Name");
			formItems.Add("Bound", true);
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			formItems.Add("Left", left_e);
			formItems.Add("Width", width_e);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("DisplayDesc", true);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			top = top + height + 1;

			formItems = new Dictionary<string, object>();
			itemName = "UomCodeS"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
			formItems.Add("Left", left_s);
			formItems.Add("Width", width_s);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("Caption", BDOSResources.getTranslate("UnitOfMeasurement"));
			formItems.Add("LinkTo", "UomCodeE");

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "UomCodeE"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "@BDOSFLTP");
			formItems.Add("Alias", "U_UomCode");
			formItems.Add("Bound", true);
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			formItems.Add("Left", left_e);
			formItems.Add("Width", width_e);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			top = top + height + 1;

			formItems = new Dictionary<string, object>();
			itemName = "PerKmS"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
			formItems.Add("Left", left_s);
			formItems.Add("Width", width_s);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("Caption", BDOSResources.getTranslate("PerHunKm"));
			formItems.Add("LinkTo", "PerKmE");

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "PerKmE"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "@BDOSFLTP");
			formItems.Add("Alias", "U_PerKm");
			formItems.Add("Bound", true);
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			formItems.Add("Left", left_e);
			formItems.Add("Width", width_e);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("DisplayDesc", true);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			top = top + height + 1;

			formItems = new Dictionary<string, object>();
			itemName = "PerHrS"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
			formItems.Add("Left", left_s);
			formItems.Add("Width", width_s);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("Caption", BDOSResources.getTranslate("PerHour"));
			formItems.Add("LinkTo", "PerHrE");

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "PerHrE"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "@BDOSFLTP");
			formItems.Add("Alias", "U_PerHr");
			formItems.Add("Bound", true);
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			formItems.Add("Left", left_e);
			formItems.Add("Width", width_e);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("DisplayDesc", true);

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

			try
			{	
				oForm.Items.Item("UomCodeE").Enabled = false;
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
		public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, out string errorText)
		{
			errorText = null;

			try
			{
				string sCFL_ID = oCFLEvento.ChooseFromListUID;
				SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

				SAPbouiCOM.DataTable oDataTable = null;
				oDataTable = oCFLEvento.SelectedObjects;

				if (oDataTable != null)
				{
					if (sCFL_ID == "StockItems_CFL")
					{
						string fuelCode = Convert.ToString(oDataTable.GetValue("ItemCode", 0));
						string itemName = Convert.ToString(oDataTable.GetValue("ItemName", 0));

						SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

						string query = @"SELECT ""UomCode"", ""ItemCode"" FROM ""OUOM"" JOIN ""OITM"" ON ""OUOM"".""UomEntry"" = ""OITM"".""IUoMEntry""  WHERE ""OITM"".""ItemCode"" = '" + fuelCode + @"'";
						oRecordSet.DoQuery(query);
						string uomCode = oRecordSet.Fields.Item("UomCode").Value;

						try
						{
							SAPbouiCOM.EditText oFuelType = (SAPbouiCOM.EditText)oForm.Items.Item("FuelTypeE").Specific;
							oFuelType.Value = fuelCode;
						}

						catch (Exception ex)
						{
							errorText = ex.Message;
						}

						try
						{
							SAPbouiCOM.EditText oItemName = (SAPbouiCOM.EditText)oForm.Items.Item("ItemNameE").Specific;
							oItemName.Value = itemName;
						}

						catch (Exception ex)
						{
							errorText = ex.Message;
						}

						try
						{
							SAPbouiCOM.EditText oUomCode = (SAPbouiCOM.EditText)oForm.Items.Item("UomCodeE").Specific;
							oUomCode.Value = uomCode;
						}

						catch (Exception ex)
						{
							errorText = ex.Message;
						}

						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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
		public static void setSizeForm(SAPbouiCOM.Form oForm, out string errorText)
		{
			errorText = null;

			try
			{
				oForm.ClientHeight = Program.uiApp.Desktop.Width / 6;
				oForm.Height = Program.uiApp.Desktop.Width / 4;
				oForm.Left = (Program.uiApp.Desktop.Width - oForm.Width) / 2;
				oForm.Top = (Program.uiApp.Desktop.Height - oForm.Height) / 3;
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
		public static void resizeForm(SAPbouiCOM.Form oForm, out string errorText)
		{
			errorText = null;

			SAPbouiCOM.Item oItem = null;

			try
			{
				oItem = oForm.Items.Item("1");
				oItem.Top = oForm.ClientHeight - 25;

				oItem = oForm.Items.Item("2");
				oItem.Top = oForm.ClientHeight - 25;
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
		public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
		{
			BubbleEvent = true;
			string errorText = null;

			if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
			{
				SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
				{
					createFormItems(oForm, out errorText);
					Program.FORM_LOAD_FOR_VISIBLE = true;
					Program.FORM_LOAD_FOR_ACTIVATE = true;
				}

				if (pVal.ItemUID == "FuelTypeE" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
				{
					if (pVal.BeforeAction == false)
					{
						SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
						oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

						BDOSConsumptionTypes.chooseFromList(oForm, oCFLEvento, out errorText);
					}
				}

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false)
				{
					if (Program.FORM_LOAD_FOR_VISIBLE == true)
					{
						oForm.Freeze(true);
						setSizeForm(oForm, out errorText);
						setVisibleFormItems(oForm, out errorText);
						oForm.Title = BDOSResources.getTranslate("ConsumTypesMasterData");
						oForm.Freeze(false);
						Program.FORM_LOAD_FOR_VISIBLE = false;
					}
					setVisibleFormItems(oForm, out errorText);
				}

				//if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false & Program.FORM_LOAD_FOR_VISIBLE == true)
				//{
				//	oForm.Freeze(true);
				//	BDOSConsumptionTypes.setSizeForm(oForm, out errorText);
				//	BDOSConsumptionTypes.setVisibleFormItems(oForm, out errorText);
				//	oForm.Title = BDOSResources.getTranslate("ConsumTypesMasterData");
				//	oForm.Freeze(false);
				//	Program.FORM_LOAD_FOR_VISIBLE = false;
				//}

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
				{
					oForm.Freeze(true);
					resizeForm(oForm, out errorText);
					oForm.Freeze(false);
				}

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE & pVal.BeforeAction == false)
				{
					if (Program.FORM_LOAD_FOR_ACTIVATE == true)
					{
						oForm.Freeze(true);
						setVisibleFormItems(oForm, out errorText);
						oForm.Freeze(false);
						//oForm.Update();
						Program.FORM_LOAD_FOR_ACTIVATE = false;
					}
				}
			}
		}
		public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
		{
			BubbleEvent = true;
			string errorText = null;

			SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

			if (oForm.TypeEx == "UDO_F_BDOSFLTP_T")
			{
				if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
				{
					setVisibleFormItems(oForm, out errorText);
				}

			}
		}
	}
}

