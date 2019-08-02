using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
	class BDOSFuelTransferWizard
	{
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
					oCreationPackage.UniqueID = "BDOSFUTRWI";
					oCreationPackage.String = BDOSResources.getTranslate("FuelTransferWizard");
					oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

					menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
				}
				catch (Exception ex)
				{
					errorText = ex.Message;
				}
			}
		}
		public static void createForm(out string errorText)
		{
			errorText = null;
			Dictionary<string, object> formItems;
			string itemName;
			SAPbouiCOM.Columns oColumns;
			SAPbouiCOM.Column oColumn;

			SAPbouiCOM.DataTable oDataTable;

			bool multiSelection;

			int top = 10;
			int height = 15;
			int left_s = 5;
			int left_e = 160;
			int width = 150;
			int left_s1 = 340;
			int left_e1 = 500;

			//ფორმის აუცილებელი თვისებები
			Dictionary<string, object> formProperties = new Dictionary<string, object>();
			formProperties.Add("UniqueID", "BDOSFUTRWI");
			formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
			formProperties.Add("Title", BDOSResources.getTranslate("FuelTransferWizard"));
			formProperties.Add("Left", 558);
			formProperties.Add("ClientWidth", 800);
			formProperties.Add("Top", 335);
			formProperties.Add("ClientHeight", 600);

			SAPbouiCOM.Form oForm;
			bool newForm;
			bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

			if (formExist == true)
			{
				if (newForm)
				{
					formItems = new Dictionary<string, object>();
					itemName = "PostDateS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("PostingDate"));
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "PostDateE";
					formItems.Add("isDataSource", true);
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("Length", 1);
					formItems.Add("Size", 20);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
					formItems.Add("TableName", "");
					formItems.Add("Alias", itemName);
					formItems.Add("Bound", true);
					formItems.Add("Left", left_e);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					top = top + height + 1;

					multiSelection = false;
					string uniqueID_WareHouse_CFL = "WareHouse_CFL";
					string objectType = "64";
					FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_WareHouse_CFL);

					formItems = new Dictionary<string, object>();
					itemName = "FromWareS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("WhsFr"));
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "FromWareE"; //10 characters
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("TableName", "");
					formItems.Add("Length", 20);
					formItems.Add("Size", 20);
					formItems.Add("Alias", itemName);
					formItems.Add("Bound", true);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
					formItems.Add("Left", left_e);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("DisplayDesc", true);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);
					formItems.Add("ChooseFromListUID", uniqueID_WareHouse_CFL);
					formItems.Add("ChooseFromListAlias", "WhsCode");

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "FromWareLB"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
					formItems.Add("Left", left_e - 20);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("LinkTo", "FromWareE");
					formItems.Add("LinkedObjectType", objectType);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					top = top + height + 1;

					multiSelection = false;
					string uniqueID_WareHouseTo_CFL = "WareHouseTo_CFL";
					objectType = "64";
					FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_WareHouseTo_CFL);

					SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_WareHouseTo_CFL);
					SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
					SAPbouiCOM.Condition oCon = oCons.Add();
					oCon.Alias = "U_BDOSWhType";
					oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
					oCon.CondVal = "Fuel";
					oCFL.SetConditions(oCons);

					formItems = new Dictionary<string, object>();
					itemName = "ToWareS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("WhsTo"));
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "ToWareE"; //10 characters
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("TableName", "");
					formItems.Add("Length", 20);
					formItems.Add("Size", 20);
					formItems.Add("Alias", itemName);
					formItems.Add("Bound", true);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
					formItems.Add("Left", left_e);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("DisplayDesc", true);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);
					formItems.Add("ChooseFromListUID", uniqueID_WareHouseTo_CFL);
					formItems.Add("ChooseFromListAlias", "WhsCode");

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "ToWareLB"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
					formItems.Add("Left", left_e - 20);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("LinkTo", "ToWareE");
					formItems.Add("LinkedObjectType", objectType);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					top = top + height + 1;

					multiSelection = false;
					string uniqueID_Project_CFL = "Project_CFL";
					objectType = "63";
					FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_Project_CFL);

					formItems = new Dictionary<string, object>();
					itemName = "FromPrjS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("FromProject"));
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "FromPrjE"; //10 characters
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("TableName", "");
					formItems.Add("Length", 20);
					formItems.Add("Size", 20);
					formItems.Add("Alias", itemName);
					formItems.Add("Bound", true);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
					formItems.Add("Left", left_e);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("DisplayDesc", true);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);
					formItems.Add("ChooseFromListUID", uniqueID_Project_CFL);
					formItems.Add("ChooseFromListAlias", "PrjCode");

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					top = top + height + 1;

					multiSelection = false;
					string uniqueID_ProjectTo_CFL = "ProjectTo_CFL";
					objectType = "63";
					FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_ProjectTo_CFL);

					formItems = new Dictionary<string, object>();
					itemName = "ToPrjS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("ToProject"));
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "ToPrjE"; //10 characters
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("TableName", "");
					formItems.Add("Length", 20);
					formItems.Add("Size", 20);
					formItems.Add("Alias", itemName);
					formItems.Add("Bound", true);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
					formItems.Add("Left", left_e);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("DisplayDesc", true);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);
					formItems.Add("ChooseFromListUID", uniqueID_ProjectTo_CFL);
					formItems.Add("ChooseFromListAlias", "PrjCode");

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					top = 10;

					multiSelection = false;
					objectType = "UDO_F_BDOSFLTP_T";
					string uniqueID_ConsumTypeCFL = "ConsumType_CFL";
					FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_ConsumTypeCFL);

					formItems = new Dictionary<string, object>();
					itemName = "ConsTypeS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s1);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("ConsumType"));
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "ConsTypeE"; //10 characters
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("TableName", "");
					formItems.Add("Length", 20);
					formItems.Add("Size", 20);
					formItems.Add("Alias", itemName);
					formItems.Add("Bound", true);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
					formItems.Add("Left", left_e1);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("DisplayDesc", true);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);
					formItems.Add("ChooseFromListUID", uniqueID_ConsumTypeCFL);
					formItems.Add("ChooseFromListAlias", "Code");

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "ConsTypeLB"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
					formItems.Add("Left", left_e1 - 20);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("LinkTo", "ConsTypeE");
					formItems.Add("LinkedObjectType", objectType);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					top = top + height + 1;

					multiSelection = false;
					objectType = "4";
					string uniqueID_VehicleCFL = "Vehicle_CFL";
					FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_VehicleCFL);

					formItems = new Dictionary<string, object>();
					itemName = "VehicleS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s1);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("Vehicle"));
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "VehicleE"; //10 characters
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("TableName", "");
					formItems.Add("Length", 20);
					formItems.Add("Size", 20);
					formItems.Add("Alias", itemName);
					formItems.Add("Bound", true);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
					formItems.Add("Left", left_e1);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("DisplayDesc", true);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);
					formItems.Add("ChooseFromListUID", uniqueID_VehicleCFL);
					formItems.Add("ChooseFromListAlias", "ItemCode");

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "VehicleLB"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
					formItems.Add("Left", left_e1 - 20);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("LinkTo", "VehicleE");
					formItems.Add("LinkedObjectType", objectType);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					top = top + height + 1;

					multiSelection = false;
					objectType = "171";
					string uniqueID_EmployeeCFL = "Employee_CFL";
					FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_EmployeeCFL);

					formItems = new Dictionary<string, object>();
					itemName = "EmployeeS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s1);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("Employee"));
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "EmployeeE"; //10 characters
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("TableName", "");
					formItems.Add("Length", 20);
					formItems.Add("Size", 20);
					formItems.Add("Alias", itemName);
					formItems.Add("Bound", true);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
					formItems.Add("Left", left_e1);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("DisplayDesc", true);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);
					formItems.Add("ChooseFromListUID", uniqueID_EmployeeCFL);
					formItems.Add("ChooseFromListAlias", "empID");

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "EmpLB"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
					formItems.Add("Left", left_e1 - 20);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("LinkTo", "EmployeeE");
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
					itemName = "FuelGroupS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s1);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("FuelGroup"));
					formItems.Add("LinkTo", "FuelGroupE");
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					Dictionary<string, string> listValidValuesItemGroups = getItemGroupsList(out errorText);

					formItems = new Dictionary<string, object>();
					itemName = "FuelGroupE";
					formItems.Add("Size", 20);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
					formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
					formItems.Add("DisplayDesc", true);
					formItems.Add("Left", left_e1);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("ValidValues", listValidValuesItemGroups);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					top = top + height + 1;

					multiSelection = false;
					objectType = "4";
					string uniqueID_FuelTyCFL = "FuelTy_CFL";
					FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_FuelTyCFL);

					//პირობის დადება აითემის არჩევის სიაზე

					oCFL = oForm.ChooseFromLists.Item(uniqueID_FuelTyCFL);
					oCons = oCFL.GetConditions();
					oCon = oCons.Add();
					oCon.Alias = "InvntItem";
					oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
					oCon.CondVal = "Y";
					oCFL.SetConditions(oCons);

					formItems = new Dictionary<string, object>();
					itemName = "FuelTypeS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s1);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("FuelType"));
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "FuelTypeE"; //10 characters
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("TableName", "");
					formItems.Add("Length", 20);
					formItems.Add("Size", 20);
					formItems.Add("Alias", itemName);
					formItems.Add("Bound", true);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
					formItems.Add("Left", left_e1);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("DisplayDesc", true);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);
					formItems.Add("ChooseFromListUID", uniqueID_FuelTyCFL);
					formItems.Add("ChooseFromListAlias", "ItemCode");

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "FuelTypeLB"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
					formItems.Add("Left", left_e1 - 20);
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

					top = top + height + 10;

					formItems = new Dictionary<string, object>();
					itemName = "Check";
					formItems.Add("Size", 20);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
					formItems.Add("Left", left_s);
					formItems.Add("Width", 19);
					formItems.Add("Top", top);
					formItems.Add("Height", 19);
					formItems.Add("UID", itemName);
					formItems.Add("Image", "HANA_CHECKBOX_CH");
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "Uncheck";
					formItems.Add("Size", 20);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
					formItems.Add("Left", left_s + 21);
					formItems.Add("Width", 19);
					formItems.Add("Top", top);
					formItems.Add("Height", 19);
					formItems.Add("UID", itemName);
					formItems.Add("Image", "HANA_CHECKBOX_UH");
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					//top = top + height + 10;
					//left_s = 5;

					//საკონტროლო პანელი
					formItems = new Dictionary<string, object>();
					itemName = "FillMTR"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
					formItems.Add("Left", left_s + 42);
					formItems.Add("Width", 100);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("Fill"));

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					top = top + height + 10;

					formItems = new Dictionary<string, object>();
					itemName = "FuelMTR"; //10 characters
					formItems.Add("isDataSource", true);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
					formItems.Add("Left", left_s);
					formItems.Add("Width", 600);
					formItems.Add("Top", top);
					formItems.Add("Height", 550);
					formItems.Add("UID", itemName);
					formItems.Add("DisplayDesc", true);
					formItems.Add("AffectsFormMode", false);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("FuelMTR").Specific;

					oColumns = oMatrix.Columns;

					SAPbouiCOM.LinkedButton oLink;

					oDataTable = oForm.DataSources.DataTables.Add("FuelMTR");

					oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ინდექსი 
					oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1);
					oDataTable.Columns.Add("AssetCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("AssetName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("Employee", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("ConsumType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("UomEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("FuelType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity, 50);
					oDataTable.Columns.Add("Dimension1", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("Dimension2", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("Dimension3", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("Dimension4", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("Dimension5", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);

					for (int count = 0; count < oDataTable.Columns.Count; count++)
					{
						var column = oDataTable.Columns.Item(count);
						string columnName = column.Name;

						if (columnName == "LineNum")
						{
							oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
							oColumn.TitleObject.Caption = "#";
							oColumn.Editable = false;
							oColumn.DataBind.Bind("FuelMTR", columnName);
							oColumn.AffectsFormMode = false;
						}
						else if (columnName == "CheckBox")
						{
							oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
							oColumn.TitleObject.Caption = "";
							oColumn.Editable = true;
							oColumn.ValOff = "N";
							oColumn.ValOn = "Y";
							oColumn.DataBind.Bind("FuelMTR", columnName);
							oColumn.AffectsFormMode = false;
						}
						else if (columnName == "AssetCode")
						{
							oColumn = oColumns.Add("AssetCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
							oColumn.TitleObject.Caption = BDOSResources.getTranslate("AssetCode");
							oColumn.Editable = false;
							oColumn.DataBind.Bind("FuelMTR", columnName);
							oLink = oColumn.ExtendedObject;
							oLink.LinkedObjectType = "4";
						}
						else if (columnName == "FuelType")
						{
							oColumn = oColumns.Add("FuelType", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
							oColumn.TitleObject.Caption = BDOSResources.getTranslate("FuelType");
							oColumn.Editable = false;
							oColumn.DataBind.Bind("FuelMTR", columnName);
							oLink = oColumn.ExtendedObject;
							oLink.LinkedObjectType = "4";
						}

						else if (columnName == "Quantity")
						{
							oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
							oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
							oColumn.Editable = true;
							oColumn.DataBind.Bind("FuelMTR", columnName);
							oColumn.AffectsFormMode = false;
						}

						else
						{
							oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
							oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
							oColumn.Editable = false;
							oColumn.DataBind.Bind("FuelMTR", columnName);
							oColumn.AffectsFormMode = false;
						}
					}

					top = oForm.Height - 70;

					//საკონტროლო პანელი
					formItems = new Dictionary<string, object>();
					itemName = "CreateDoc"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
					formItems.Add("Left", left_s);
					formItems.Add("Width", 150);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("CreateDocument"));

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					oMatrix.Clear();
					oMatrix.LoadFromDataSource();
					oMatrix.AutoResizeColumns();
				}
				resizeItems(oForm);
				oForm.Visible = true;
				oForm.Select();
			}
		}
		private static void CreateStockTransfer(SAPbouiCOM.Form oForm)
		{
			int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreateStockTransfer") + " ?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

			if (answer == 2)
			{
				return;
			}

			SAPbouiCOM.EditText oEditTextDocDate = (SAPbouiCOM.EditText)oForm.Items.Item("PostDateE").Specific;
			string DocDateS = oEditTextDocDate.Value;

			DateTime DocDate = Convert.ToDateTime(DateTime.ParseExact(DocDateS, "yyyyMMdd", CultureInfo.InvariantCulture));

			SAPbobsCOM.StockTransfer oStockTransfer = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
			SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("FuelMTR").Specific;
			oStockTransfer.DocDate = DocDate;
			oStockTransfer.TaxDate = DocDate;
			oStockTransfer.FromWarehouse = oForm.Items.Item("FromWareE").Specific.Value;
			oStockTransfer.ToWarehouse = oForm.Items.Item("ToWareE").Specific.Value;
			oStockTransfer.UserFields.Fields.Item("U_BDOSFrPrj").Value = oForm.Items.Item("FromPrjE").Specific.Value;

			int count = oMatrix.RowCount;
			double quantity;

			for (int i = 1; i <= oMatrix.RowCount; i++)
			{
				bool checkedLine = oMatrix.GetCellSpecific("CheckBox", i).Checked;

				if (checkedLine)
				{
					if (oMatrix.GetCellSpecific("FuelType", i).Value == "")
					{
						continue;
					}

					oStockTransfer.Lines.ItemCode = oMatrix.GetCellSpecific("FuelType", i).Value;
					oStockTransfer.Lines.ItemDescription = oMatrix.GetCellSpecific("AssetName", i).Value;

					if (oMatrix.GetCellSpecific("Quantity", i).Value != "")
					{
						quantity = Convert.ToDouble(oMatrix.GetCellSpecific("Quantity", i).Value);
						oStockTransfer.Lines.Quantity = quantity;
					}

					oStockTransfer.Lines.FromWarehouseCode = oForm.Items.Item("FromWareE").Specific.Value;
					oStockTransfer.Lines.WarehouseCode = oForm.Items.Item("ToWareE").Specific.Value;
					oStockTransfer.Lines.DistributionRule = oMatrix.GetCellSpecific("Dimension1", i).Value;
					oStockTransfer.Lines.DistributionRule2 = oMatrix.GetCellSpecific("Dimension2", i).Value;
					oStockTransfer.Lines.DistributionRule3 = oMatrix.GetCellSpecific("Dimension3", i).Value;
					oStockTransfer.Lines.DistributionRule4 = oMatrix.GetCellSpecific("Dimension4", i).Value;
					oStockTransfer.Lines.DistributionRule5 = oMatrix.GetCellSpecific("Dimension5", i).Value;

					oStockTransfer.Lines.Add();
				}
			}

			int resultCode = oStockTransfer.Add();

			if (resultCode != 0)
			{
				string errorMessage = "";
				Program.oCompany.GetLastError(out resultCode, out errorMessage);
				Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + errorMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
			}
			else
			{
				string docEntry;
				Program.oCompany.GetNewObjectCode(out docEntry);
				Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCreatedSuccesfully") + ": " + docEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
			}
		}
		public static void resizeItems(SAPbouiCOM.Form oForm)
		{
			try
			{
				SAPbouiCOM.Item oMatrixItem = oForm.Items.Item("FuelMTR");

				oMatrixItem.Height = oForm.Height - 220;
				oMatrixItem.Width = oForm.Width - 20;
			}
			catch
			{
			}
		}
		public static void fillMTRItems(SAPbouiCOM.Form oForm, out string errorText)
		{
			errorText = null;

			SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("FuelMTR");
			SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("FuelMTR").Specific;
			oDataTable.Rows.Clear();

			SAPbouiCOM.EditText oEditTextDate = (SAPbouiCOM.EditText)oForm.Items.Item("PostDateE").Specific;
			String postDate = oEditTextDate.Value;
			if (string.IsNullOrEmpty(postDate))
			{
				errorText = BDOSResources.getTranslate("PostingDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
				Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
				return;
			}

			SAPbouiCOM.EditText oEditTextWareFrom = (SAPbouiCOM.EditText)oForm.Items.Item("FromWareE").Specific;
			String wareFrom = oEditTextWareFrom.Value;
			if (string.IsNullOrEmpty(wareFrom))
			{
				errorText = BDOSResources.getTranslate("WhsFr") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
				Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
				return;
			}

			SAPbouiCOM.EditText oEditTextWareTo = (SAPbouiCOM.EditText)oForm.Items.Item("ToWareE").Specific;
			String wareTo = oEditTextWareTo.Value;
			if (string.IsNullOrEmpty(wareTo))
			{
				errorText = BDOSResources.getTranslate("WhsTo") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
				Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
				return;
			}

			SAPbouiCOM.EditText oEditTextFromPrj = (SAPbouiCOM.EditText)oForm.Items.Item("FromPrjE").Specific;
			String fromPrj = oEditTextFromPrj.Value;
			if (string.IsNullOrEmpty(fromPrj))
			{
				errorText = BDOSResources.getTranslate("FromProject") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
				Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
				return;
			}

			SAPbouiCOM.EditText oEditTextToPrj = (SAPbouiCOM.EditText)oForm.Items.Item("ToPrjE").Specific;
			String toPrj = oEditTextToPrj.Value;
			if (string.IsNullOrEmpty(toPrj))
			{
				errorText = BDOSResources.getTranslate("ToProject") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
				Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
				return;
			}

			string consType;
			consType = oForm.Items.Item("ConsTypeE").Specific.Value;
			bool eCons = string.IsNullOrEmpty(consType);

			string vehicle;
			vehicle = oForm.Items.Item("VehicleE").Specific.Value;
			bool eVeh = string.IsNullOrEmpty(vehicle);

			string employee;
			employee = oForm.Items.Item("EmployeeE").Specific.Value;
			bool eEmp = string.IsNullOrEmpty(employee);

			string fuelGroup;
			fuelGroup = oForm.Items.Item("FuelGroupE").Specific.Value;
			bool eFuelGr = string.IsNullOrEmpty(fuelGroup);

			string fuelType;
			fuelType = oForm.Items.Item("FuelTypeE").Specific.Value;
			bool eFuelTy = string.IsNullOrEmpty(fuelType);

			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.Recordset oRecordSetDimension = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				StringBuilder queryBuilder = new StringBuilder();

				queryBuilder.Append(
					@"SELECT 
                        ""OITM"".""ItemCode"",
                        ""OITM"".""ItemName"",
                        ""OPRC"".""PrcCode"" AS ""CostCentre"",
                        ""OHEM"".""empID"" AS ""Employee"",
                        ""OHEM"".""lastName"" AS ""LastName"",
                        ""OHEM"".""firstName"" AS ""FirstName"",
                        ""OITM"".""U_FltCode"" AS ""FuelCode"",
                        ""OITM"".""U_UomCode"" AS ""UomCode"",
                        ""OITM"".""ItmsGrpCod"" AS ""FuelGroup"",
                        ""OITM"".""U_FuelType"" AS ""FuelType""
                    FROM 
                        ""OITM""
                    LEFT JOIN 
                        ""OHEM"" ON ""OITM"".""Employee"" = ""OHEM"".""empID""
                    LEFT JOIN 
                        ""OPRC"" ON ""OITM"".""ItemCode"" = ""OPRC"".""U_BDOSFACode""
					INNER JOIN 
						""OACS"" ON ""OITM"".""AssetClass"" = ""OACS"".""Code"" AND ""OACS"".""U_visCode"" = 'Y'
					WHERE 
                        1 = 1 ");

				if (!eCons)
				{
					queryBuilder.Append(@" AND ""OITM"".""U_FltCode"" = '" + consType + @"'");
				}

				if (!eVeh)
				{
					queryBuilder.Append(@" AND ""OITM"".""ItemCode"" = '" + vehicle + @"'");
				}

				if (!eEmp)
				{
					queryBuilder.Append(@" AND ""OHEM"".""empID"" = '" + employee + @"'");
				}

				if (!eFuelGr)
				{
					queryBuilder.Append(@" AND  ""OITM"".""ItmsGrpCod"" = '" + fuelGroup + @"'");
				}

				if (!eFuelTy)
				{
					queryBuilder.Append(@" AND  ""OITM"".""U_FuelType"" = '" + fuelType + @"'");
				}

				queryBuilder.Append(@" ORDER BY ""OITM"".""ItemCode""");

				string query = queryBuilder.ToString();
				oRecordSet.DoQuery(query);

				string queryDim = @"SELECT ""U_BDOSFADim"" AS ""Dimension"" FROM ""OADM""";
				oRecordSetDimension.DoQuery(queryDim);
				var dimColumn = "Dimension" + oRecordSetDimension.Fields.Item("Dimension").Value;

				int rowIndex = 0;

				while (!oRecordSet.EoF)
				{
					oDataTable.Rows.Add();

					oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
					oDataTable.SetValue("CheckBox", rowIndex, "N");
					oDataTable.SetValue("AssetCode", rowIndex, oRecordSet.Fields.Item("ItemCode").Value);
					oDataTable.SetValue("AssetName", rowIndex, oRecordSet.Fields.Item("ItemName").Value);
					oDataTable.SetValue("Employee", rowIndex, oRecordSet.Fields.Item("LastName").Value + " " + oRecordSet.Fields.Item("FirstName").Value);
					oDataTable.SetValue("ConsumType", rowIndex, oRecordSet.Fields.Item("FuelCode").Value);
					oDataTable.SetValue("UomEntry", rowIndex, oRecordSet.Fields.Item("UomCode").Value);
					oDataTable.SetValue("FuelType", rowIndex, oRecordSet.Fields.Item("FuelType").Value);
					oDataTable.SetValue(dimColumn, rowIndex, oRecordSet.Fields.Item("CostCentre").Value);

					rowIndex++;
					oRecordSet.MoveNext();
				}
			}

			catch (Exception ex)
			{
				errorText = ex.Message;
			}

			oForm.Freeze(true);
			oMatrix.Clear();
			oMatrix.LoadFromDataSource();
			oMatrix.AutoResizeColumns();
			oForm.Update();
			oForm.Freeze(false);
		}
		public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
		{
			BubbleEvent = true;
			string errorText = null;

			if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
			{
				SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
				{
					SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
					oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

					chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
				}

				if (pVal.ItemUID == "FillMTR" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
				{
					fillMTRItems(oForm, out errorText);
				}

				if ((pVal.ItemUID == "Check" || pVal.ItemUID == "Uncheck") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
				{
					checkUncheckTables(oForm, pVal.ItemUID, "FuelMTR", out errorText);
				}

				if (pVal.ItemUID == "CreateDoc" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
				{
					CreateStockTransfer(oForm);
					fillMTRItems(oForm, out errorText);
				}
			}
		}
		private static void checkUncheckTables(SAPbouiCOM.Form oForm, string CheckOperation, string matrixName, out string errorText)
		{
			errorText = null;

			oForm.Freeze(true);

			SAPbouiCOM.CheckBox oCheckBox;
			SAPbouiCOM.Matrix oMatrix;

			oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item(matrixName).Specific));

			int rowCount = oMatrix.RowCount;
			for (int j = 1; j <= rowCount; j++)
			{
				oCheckBox = oMatrix.Columns.Item("CheckBox").Cells.Item(j).Specific;
				oCheckBox.Checked = (CheckOperation == "Check");
			}
			oForm.Freeze(false);
		}
		private static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string pVal, bool BeforeAction, out string errorText)
		{
			errorText = null;

			string sCFL_ID = oCFLEvento.ChooseFromListUID;
			SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			string query = "";

			if (BeforeAction == false)
			{
				SAPbouiCOM.DataTable oDataTable = null;
				oDataTable = oCFLEvento.SelectedObjects;

				if (oDataTable != null)
				{
					try
					{
						if (sCFL_ID == "WareHouse_CFL")
						{
							string whsCode = Convert.ToString(oDataTable.GetValue("WhsCode", 0));

							try
							{
								SAPbouiCOM.EditText oWareFrom = oForm.Items.Item("FromWareE").Specific;
								oWareFrom.Value = whsCode;

							}
							catch (Exception ex)
							{
								errorText = ex.Message;
							}

							try
							{
								query = @"SELECT ""WhsCode"",""U_BDOSPrjCod"" AS ""PrjFr"" FROM ""OWHS"" WHERE ""WhsCode"" = '" + whsCode + @"'";
								oRecordSet.DoQuery(query);
								string frValue = oRecordSet.Fields.Item("PrjFr").Value;

								SAPbouiCOM.EditText oFromPrj = oForm.Items.Item("FromPrjE").Specific;
								oFromPrj.Value = frValue;
							}

							catch (Exception ex)
							{
								errorText = ex.Message;
							}
						}

						if (sCFL_ID == "WareHouseTo_CFL")
						{
							string whsCodeTo = Convert.ToString(oDataTable.GetValue("WhsCode", 0));

							try
							{
								SAPbouiCOM.EditText oWareTo = oForm.Items.Item("ToWareE").Specific;
								oWareTo.Value = whsCodeTo;
							}

							catch (Exception ex)
							{
								errorText = ex.Message;
							}

							try
							{
								query = @"SELECT ""WhsCode"",""U_BDOSPrjCod"" AS ""PrjFr"" FROM ""OWHS"" WHERE ""WhsCode"" = '" + whsCodeTo + @"'";
								oRecordSet.DoQuery(query);
								string toValue = oRecordSet.Fields.Item("PrjFr").Value;

								SAPbouiCOM.EditText oToPrj = oForm.Items.Item("ToPrjE").Specific;
								oToPrj.Value = toValue;
							}

							catch (Exception ex)
							{
								errorText = ex.Message;
							}
						}

						if (sCFL_ID == "ConsumType_CFL")
						{
							string consCode = Convert.ToString(oDataTable.GetValue("Code", 0));

							try
							{
								SAPbouiCOM.EditText oConsType = oForm.Items.Item("ConsTypeE").Specific;
								oConsType.Value = consCode;
							}
							catch (Exception ex)
							{
								errorText = ex.Message;
							}
						}

						if (sCFL_ID == "Vehicle_CFL")
						{
							string vehCode = Convert.ToString(oDataTable.GetValue("ItemCode", 0));

							try
							{
								SAPbouiCOM.EditText oVehicle = oForm.Items.Item("VehicleE").Specific;
								oVehicle.Value = vehCode;
							}
							catch (Exception ex)
							{
								errorText = ex.Message;
							}
						}

						if (sCFL_ID == "Employee_CFL")
						{
							string empCode = Convert.ToString(oDataTable.GetValue("empID", 0));

							try
							{
								SAPbouiCOM.EditText oEmployee = oForm.Items.Item("EmployeeE").Specific;
								oEmployee.Value = empCode;
							}
							catch (Exception ex)
							{
								errorText = ex.Message;
							}
						}

						if (sCFL_ID == "FuelTy_CFL")
						{
							string fuelType = Convert.ToString(oDataTable.GetValue("ItemCode", 0));

							try
							{
								SAPbouiCOM.EditText oFuelType = oForm.Items.Item("FuelTypeE").Specific;
								oFuelType.Value = fuelType;
							}
							catch (Exception ex)
							{
								errorText = ex.Message;
							}
						}

						if (sCFL_ID == "Project_CFL")
						{
							string fromProject = Convert.ToString(oDataTable.GetValue("PrjCode", 0));

							try
							{
								SAPbouiCOM.EditText oFromPrj = oForm.Items.Item("FromPrjE").Specific;
								oFromPrj.Value = fromProject;
							}
							catch (Exception ex)
							{
								errorText = ex.Message;
							}
						}

						if (sCFL_ID == "ProjectTo_CFL")
						{
							string toProject = Convert.ToString(oDataTable.GetValue("PrjCode", 0));

							try
							{
								SAPbouiCOM.EditText oToPrj = oForm.Items.Item("ToPrjE").Specific;
								oToPrj.Value = toProject;
							}
							catch (Exception ex)
							{
								errorText = ex.Message;
							}
						}
					}
					catch (Exception ex)
					{
						errorText = ex.Message;
					}
				}
			}
			else
			{
				if (sCFL_ID == "Vehicle_CFL")
				{
					oForm.Freeze(true);
					try
					{
						string queryVis = @"SELECT ""ItemCode""
                                                    FROM ""OITM""
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
		public static Dictionary<string, string> getItemGroupsList(out string errorText)
		{
			errorText = null;

			Dictionary<string, string> grpList = new Dictionary<string, string>();
			grpList.Add("", "");
			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			string query = @"SELECT * FROM ""OITB"" WHERE ""Locked""='N'";

			oRecordSet.DoQuery(query);

			while (!oRecordSet.EoF)
			{
				grpList.Add(oRecordSet.Fields.Item("ItmsGrpCod").Value.ToString(), oRecordSet.Fields.Item("ItmsGrpNam").Value.ToString());
				oRecordSet.MoveNext();
			}

			return grpList;
		}
	}
}
