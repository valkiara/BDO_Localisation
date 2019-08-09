﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace BDO_Localisation_AddOn
{
	static partial class BDO_Vehicles
	{
		public static void createMasterDataUDO(out string errorText)
		{
			errorText = null;
			string tableName = "BDO_VECL";
			string description = "Vehicles";

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
			fieldskeysMap.Add("Name", "number");
			fieldskeysMap.Add("TableName", "BDO_VECL");
			fieldskeysMap.Add("Description", "Number");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 20);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "brand");
			fieldskeysMap.Add("TableName", "BDO_VECL");
			fieldskeysMap.Add("Description", "Brand");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 20);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "trailNum");
			fieldskeysMap.Add("TableName", "BDO_VECL");
			fieldskeysMap.Add("Description", "Trailer Number");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 20);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "drvCode");
			fieldskeysMap.Add("TableName", "BDO_VECL");
			fieldskeysMap.Add("Description", "Driver Code");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("LinkedTable", "BDO_DRVS");

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			GC.Collect();
		}

		public static void registerUDO(out string errorText)
		{
			errorText = null;
			string code = "UDO_F_BDO_VECL_D"; //20 characters (must include at least one alphabetical character).
			Dictionary<string, object> formProperties;

			formProperties = new Dictionary<string, object>();
			formProperties.Add("Name", "Vehicles"); //100 characters
			formProperties.Add("TableName", "BDO_VECL");
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
			fieldskeysMap.Add("ColumnAlias", "U_number");
			fieldskeysMap.Add("ColumnDescription", "Number"); //30 characters
			listFindColumns.Add(fieldskeysMap);
			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("ColumnAlias", "U_brand");
			fieldskeysMap.Add("ColumnDescription", "Brand"); //30 characters
			listFindColumns.Add(fieldskeysMap);
			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("ColumnAlias", "U_trailNum");
			fieldskeysMap.Add("ColumnDescription", "Trailer Number"); //30 characters
			listFindColumns.Add(fieldskeysMap);
			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("ColumnAlias", "U_drvCode");
			fieldskeysMap.Add("ColumnDescription", "Driver Code"); //30 characters
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

			try
			{
				fatherMenuItem = Program.uiApp.Menus.Item("43544");
				// Add a pop-up menu item
				oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
				oCreationPackage.Checked = false;
				oCreationPackage.Enabled = true;
				oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
				oCreationPackage.UniqueID = "UDO_F_BDO_VECL_D";
				oCreationPackage.String = BDOSResources.getTranslate("VehicleMasterData");
				oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

				menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
			}
			catch (Exception ex)
			{
				errorText = ex.Message;
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

			formItems = new Dictionary<string, object>();
			itemName = "13_U_S"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
			formItems.Add("Left", left_s);
			formItems.Add("Width", width_s);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("Caption", BDOSResources.getTranslate("Number"));
			formItems.Add("LinkTo", "13_U_E");

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "13_U_E"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "@BDO_VECL");
			formItems.Add("Alias", "U_number");
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
			itemName = "14_U_S"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
			formItems.Add("Left", left_s);
			formItems.Add("Width", width_s);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("Caption", BDOSResources.getTranslate("Brand"));
			formItems.Add("LinkTo", "14_U_E");

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "14_U_E"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "@BDO_VECL");
			formItems.Add("Alias", "U_brand");
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
			itemName = "15_U_S"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
			formItems.Add("Left", left_s);
			formItems.Add("Width", width_s);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("Caption", BDOSResources.getTranslate("TrailerNumber"));
			formItems.Add("LinkTo", "15_U_E");

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "15_U_E"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "@BDO_VECL");
			formItems.Add("Alias", "U_trailNum");
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

			bool multiSelection = false;
			string objectType = "UDO_F_BDO_DRVS_D";
			string uniqueID_CFL = "DriverCode_CFL";
			FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_CFL);

			formItems = new Dictionary<string, object>();
			itemName = "16_U_S"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
			formItems.Add("Left", left_s);
			formItems.Add("Width", width_s);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("Caption", BDOSResources.getTranslate("DriverCode"));
			formItems.Add("LinkTo", "15_U_E");

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "16_U_E"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "@BDO_VECL");
			formItems.Add("Alias", "U_drvCode");
			formItems.Add("Bound", true);
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			formItems.Add("Left", left_e);
			formItems.Add("Width", width_e);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("DisplayDesc", true);
			formItems.Add("ChooseFromListUID", uniqueID_CFL);
			formItems.Add("ChooseFromListAlias", "Code");

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "16_U_LB"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
			formItems.Add("Left", left_e - 20);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("LinkTo", "16_U_E");
			formItems.Add("LinkedObjectType", objectType);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			GC.Collect();
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
					string driverCode = Convert.ToString(oDataTable.GetValue("Code", 0));
					oForm.DataSources.DBDataSources.Item("@BDO_VECL").SetValue("U_drvCode", 0, driverCode);

					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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
				//oForm.ClientWidth = Program.uiApp.Desktop.Width / 3;

				oForm.Height = Program.uiApp.Desktop.Width / 4;
				//oForm.ClientWidth = Program.uiApp.Desktop.Width / 2;

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

		//public static void formDataLoad(  SAPbouiCOM.Form oForm, out string errorText)
		//{
		//    errorText = null;

		//    BDOSResources.getTranslate("VehicleMasterData");
		//}

		public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
		{
			BubbleEvent = true;
			string errorText = null;

			if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
			{
				SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == false)
				{
					BDO_Vehicles.createFormItems(oForm, out errorText);
					Program.FORM_LOAD_FOR_VISIBLE = true;
				}

				if (pVal.ItemUID == "16_U_E" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
				{
					if (pVal.BeforeAction == false)
					{
						SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
						oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

						BDO_Vehicles.chooseFromList(oForm, oCFLEvento, out errorText);
					}
				}

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false & Program.FORM_LOAD_FOR_VISIBLE == true)
				{
					oForm.Freeze(true);
					BDO_Vehicles.setSizeForm(oForm, out errorText);
					oForm.Title = BDOSResources.getTranslate("VehicleMasterData");
					oForm.Freeze(false);
					Program.FORM_LOAD_FOR_VISIBLE = false;
				}

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
				{
					oForm.Freeze(true);
					BDO_Vehicles.resizeForm(oForm, out errorText);
					oForm.Freeze(false);
				}
			}
		}

		public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
		{
			BubbleEvent = true;
			string errorText = null;

			SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

			if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE)
			{
				if (BusinessObjectInfo.BeforeAction == true)  //& pVal.InnerEvent == true)
				{
					if (checkRemoving(oForm, out errorText) == true)
					{
						Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("RecordIsUsedInDocuments"));
						Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
						BubbleEvent = false;
					}
				}
			}
		}

		public static bool checkRemoving(SAPbouiCOM.Form oForm, out string errorText)
		{
			errorText = null;

			SAPbouiCOM.DBDataSource DocDBSourceTAXP = oForm.DataSources.DBDataSources.Item(0);
			string code = DocDBSourceTAXP.GetValue("Code", 0).Trim();

			Dictionary<string, string> listTables = new Dictionary<string, string>();
			listTables.Add("@BDO_WBLD", "U_vehicle"); //Waybills

			return CommonFunctions.codeIsUsed(listTables, code);
		}
	}
}
