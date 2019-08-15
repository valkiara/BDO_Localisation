using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
	class FixedAsset
	{
		public static SAPbouiCOM.Form CurrentForm;
		public static void createUserFields(out string errorText)
		{
			errorText = null;
			Dictionary<string, object> fieldskeysMap;

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "BDOSCostAc");
			fieldskeysMap.Add("TableName", "OITM");
			fieldskeysMap.Add("Description", "Cost accounting object");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 150);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "BDOSFACode");
			fieldskeysMap.Add("TableName", "OPRC");
			fieldskeysMap.Add("Description", "Fixed asset code");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 150);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "FltCode");
			fieldskeysMap.Add("TableName", "OITM");
			fieldskeysMap.Add("Description", "Code");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 50);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "FuelType");
			fieldskeysMap.Add("TableName", "OITM");
			fieldskeysMap.Add("Description", "Fuel Type");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 50);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "Name");
			fieldskeysMap.Add("TableName", "OITM");
			fieldskeysMap.Add("Description", "Name");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 100);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "UomCode");
			fieldskeysMap.Add("TableName", "OITM");
			fieldskeysMap.Add("Description", "Uom Code");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 20);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "PerKm"); //
			fieldskeysMap.Add("TableName", "OITM");
			fieldskeysMap.Add("Description", "Per 100 km");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
			fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Quantity);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "PerHr"); //
			fieldskeysMap.Add("TableName", "OITM");
			fieldskeysMap.Add("Description", "Per hour");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
			fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Quantity);

			UDO.addUserTableFields(fieldskeysMap, out errorText);


			GC.Collect();

		}

		public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
		{
			errorText = null;

			Dictionary<string, object> formItems;
			string itemName = "";

			int left_s = oForm.Items.Item("1470002161").Left;
			int left_e = oForm.Items.Item("1470002162").Left;
			int height = oForm.Items.Item("1470002161").Height;
			int top = oForm.Items.Item("1470002316").Top;
			int width_s = oForm.Items.Item("1470002161").Width;
			int width_e = oForm.Items.Item("1470002162").Width;

			int pane = 11;

			formItems = new Dictionary<string, object>();
			itemName = "Header"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
			formItems.Add("Left", left_s);
			formItems.Add("Width", width_s);
			formItems.Add("Top", top - height);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("Caption", BDOSResources.getTranslate("FuelMng"));
			formItems.Add("FromPane", pane);
			formItems.Add("ToPane", pane);
			formItems.Add("Visible", false);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			top = top + height + 1;

			bool multiSelection = false;
			string objectType = "UDO_F_BDOSFLTP_T";
			string uniqueID_AssetClassCFL = "Assets_CFL";
			FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_AssetClassCFL);

			formItems = new Dictionary<string, object>();
			itemName = "FltCodeS"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
			formItems.Add("Left", left_s);
			formItems.Add("Width", width_s);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("Caption", BDOSResources.getTranslate("Code"));
			formItems.Add("FromPane", pane);
			formItems.Add("ToPane", pane);
			formItems.Add("LinkTo", "FltCodeE");
			formItems.Add("Visible", false);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "FltCodeE"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "OITM");
			formItems.Add("Alias", "U_FltCode");
			formItems.Add("Bound", true);
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			formItems.Add("Left", left_e);
			formItems.Add("Width", width_e);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("DisplayDesc", true);
			formItems.Add("FromPane", pane);
			formItems.Add("ToPane", pane);
			formItems.Add("ChooseFromListUID", uniqueID_AssetClassCFL);
			formItems.Add("ChooseFromListAlias", "Code");
			formItems.Add("Visible", false);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "FltCodeLB"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
			formItems.Add("Left", left_e - 20);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("LinkTo", "FltCodeE");
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
			itemName = "FuelTypeS"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
			formItems.Add("Left", left_s);
			formItems.Add("Width", width_s);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("Caption", BDOSResources.getTranslate("FuelType"));
			formItems.Add("FromPane", pane);
			formItems.Add("ToPane", pane);
			formItems.Add("LinkTo", "FuelTypeE");
			formItems.Add("Visible", false);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "FuelTypeE"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "OITM");
			formItems.Add("Alias", "U_FuelType");
			formItems.Add("Bound", true);
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			formItems.Add("Left", left_e);
			formItems.Add("Width", width_e);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("DisplayDesc", true);
			formItems.Add("FromPane", pane);
			formItems.Add("ToPane", pane);
			formItems.Add("Visible", false);

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
			formItems.Add("FromPane", pane);
			formItems.Add("ToPane", pane);
			formItems.Add("LinkTo", "ItemNameE");
			formItems.Add("Visible", false);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "ItemNameE"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "OITM");
			formItems.Add("Alias", "U_Name");
			formItems.Add("Bound", true);
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			formItems.Add("Left", left_e);
			formItems.Add("Width", width_e);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("DisplayDesc", true);
			formItems.Add("FromPane", pane);
			formItems.Add("ToPane", pane);
			formItems.Add("Visible", false);

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
			formItems.Add("FromPane", pane);
			formItems.Add("ToPane", pane);
			formItems.Add("LinkTo", "UomCodeE");
			formItems.Add("Visible", false);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "UomCodeE"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "OITM");
			formItems.Add("Alias", "U_UomCode");
			formItems.Add("Bound", true);
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			formItems.Add("Left", left_e);
			formItems.Add("Width", width_e);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("DisplayDesc", true);
			formItems.Add("FromPane", pane);
			formItems.Add("ToPane", pane);
			formItems.Add("Visible", false);

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
			formItems.Add("FromPane", pane);
			formItems.Add("ToPane", pane);
			formItems.Add("LinkTo", "PerKmE");
			formItems.Add("Visible", false);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "PerKmE"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "OITM");
			formItems.Add("Alias", "U_PerKm");
			formItems.Add("Bound", true);
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			formItems.Add("Left", left_e);
			formItems.Add("Width", width_e);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("DisplayDesc", true);
			formItems.Add("FromPane", pane);
			formItems.Add("ToPane", pane);
			formItems.Add("Visible", false);

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
			formItems.Add("FromPane", pane);
			formItems.Add("ToPane", pane);
			formItems.Add("LinkTo", "PerHrE");
			formItems.Add("Visible", false);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			formItems = new Dictionary<string, object>();
			itemName = "PerHrE"; //10 characters
			formItems.Add("isDataSource", true);
			formItems.Add("DataSource", "DBDataSources");
			formItems.Add("TableName", "OITM");
			formItems.Add("Alias", "U_PerHr");
			formItems.Add("Bound", true);
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			formItems.Add("Left", left_e);
			formItems.Add("Width", width_e);
			formItems.Add("Top", top);
			formItems.Add("Height", height);
			formItems.Add("UID", itemName);
			formItems.Add("DisplayDesc", true);
			formItems.Add("FromPane", pane);
			formItems.Add("ToPane", pane);
			formItems.Add("Visible", false);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			SAPbouiCOM.Item oItem = oForm.Items.Item("FuelTypeS");
			top = oItem.Top;

			string caption = BDOSResources.getTranslate("CreateDistributionRule");
			formItems = new Dictionary<string, object>();
			itemName = "BDODistTXT"; //10 characters
			formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
			formItems.Add("Left", oForm.Items.Item("122").Left);
			formItems.Add("Width", oForm.Items.Item("122").Width);
			formItems.Add("Top", oForm.Items.Item("122").Top + 20);
			formItems.Add("Height", 14);
			formItems.Add("UID", itemName);
			formItems.Add("Caption", caption);
			formItems.Add("TextStyle", 4);
			formItems.Add("FontSize", 10);
			formItems.Add("Enabled", true);
			formItems.Add("FromPane", 6);
			formItems.Add("ToPane", 6);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}

			oForm.DataSources.UserDataSources.Add("BDSDistCod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

			//string objectType = "87"; //DistRule
			//string uniqueID_DistCFL = "Dist_CFL";
			//FormsB1.addChooseFromList(oForm, false, objectType, uniqueID_DistCFL);

			//formItems = new Dictionary<string, object>();
			//itemName = "BDODistCod"; //10 characters
			//formItems.Add("isDataSource", true);
			//formItems.Add("DataSource", "UserDataSources");
			//formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
			//formItems.Add("Length", 11);
			//formItems.Add("TableName", "");
			//formItems.Add("Alias", itemName);
			//formItems.Add("Bound", true);
			//formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
			//formItems.Add("Left", 250);//oForm.Items.Item("122").Left + oForm.Items.Item("122").Width + 40);
			//formItems.Add("Width", 50);
			//formItems.Add("Top", oForm.Items.Item("122").Top + 20);
			//formItems.Add("Height", 14);
			//formItems.Add("UID", itemName);
			//formItems.Add("AffectsFormMode", false);
			//formItems.Add("DisplayDesc", true);
			//formItems.Add("Enabled", false);
			//formItems.Add("ChooseFromListUID", uniqueID_DistCFL);
			//formItems.Add("ChooseFromListAlias", "Code");
			//formItems.Add("FromPane", 6);
			//formItems.Add("ToPane", 6);

			//FormsB1.createFormItem(oForm, formItems, out errorText);
			//if (errorText != null)
			//{
			//    return;
			//}

			//formItems = new Dictionary<string, object>();
			//itemName = "BDODistLB"; //10 characters
			//formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
			//formItems.Add("Left", oForm.Items.Item("122").Left + oForm.Items.Item("122").Width + 40 - 20);
			//formItems.Add("Top", oForm.Items.Item("122").Top + 20);
			//formItems.Add("Height", 14);
			//formItems.Add("Width", 15);
			//formItems.Add("UID", itemName);
			//formItems.Add("LinkTo", "BDODistCod");
			//formItems.Add("LinkedObjectType", objectType);
			//formItems.Add("FromPane", 6);
			//formItems.Add("ToPane", 6);

			FormsB1.createFormItem(oForm, formItems, out errorText);
			if (errorText != null)
			{
				return;
			}
		}
		public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
		{
			errorText = null;

			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			string itemValue = oForm.Items.Item("1470002152").Specific.Value;
			
			string query =
				@"SELECT DISTINCT
					""OACS"".""Code"" AS ""AssetClass""
				FROM 
					""OITM""
				INNER JOIN 
					""OACS"" ON ""OITM"".""AssetClass"" = ""OACS"".""Code"" AND ""OACS"".""U_visCode"" = 'Y'
				WHERE 
					""OACS"".""Code"" = '" + itemValue + @"'";

			try
			{
				oRecordSet.DoQuery(query);

				if (!oRecordSet.EoF && oForm.PaneLevel == 11)
				{
					oForm.Items.Item("Header").Visible = true;
					oForm.Items.Item("FltCodeS").Visible = true;
					oForm.Items.Item("FltCodeE").Visible = true;
					oForm.Items.Item("FltCodeLB").Visible = true;
					oForm.Items.Item("FuelTypeS").Visible = true;
					oForm.Items.Item("FuelTypeE").Visible = true;
					oForm.Items.Item("ItemNameS").Visible = true;
					oForm.Items.Item("ItemNameE").Visible = true;
					oForm.Items.Item("UomCodeS").Visible = true;
					oForm.Items.Item("UomCodeE").Visible = true;
					oForm.Items.Item("PerKmS").Visible = true;
					oForm.Items.Item("PerKmE").Visible = true;
					oForm.Items.Item("PerHrS").Visible = true;
					oForm.Items.Item("PerHrE").Visible = true;
				}

				else
				{
					oForm.Items.Item("Header").Visible = false;
					oForm.Items.Item("FltCodeS").Visible = false;
					oForm.Items.Item("FltCodeE").Visible = false;
					oForm.Items.Item("FltCodeLB").Visible = false;
					oForm.Items.Item("FuelTypeS").Visible = false;
					oForm.Items.Item("FuelTypeE").Visible = false;
					oForm.Items.Item("ItemNameS").Visible = false;
					oForm.Items.Item("ItemNameE").Visible = false;
					oForm.Items.Item("UomCodeS").Visible = false;
					oForm.Items.Item("UomCodeE").Visible = false;
					oForm.Items.Item("PerKmS").Visible = false;
					oForm.Items.Item("PerKmE").Visible = false;
					oForm.Items.Item("PerHrS").Visible = false;
					oForm.Items.Item("PerHrE").Visible = false;
				}
			}

			catch (Exception ex)
			{
				errorText = ex.Message;
			}
			finally
			{
				oForm.Update();
				GC.Collect();
			}
		}
		public static void fillItems(SAPbouiCOM.Form oForm, out string errorText)
		{
			errorText = null;

			StringBuilder queryBuilder = new StringBuilder();
			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			string itemValue = oForm.Items.Item("1470002152").Specific.Value;

			queryBuilder.Append(
				@"SELECT DISTINCT
					""OACS"".""Code"" AS ""AssetClass"",
					""OACS"".""U_Code"" AS ""AssetConsType"",
					""@BDOSFLTP"".""Code"" AS ""ConsType"",
					""@BDOSFLTP"".""U_Name"" AS ""Name"",
					""@BDOSFLTP"".""U_FuelType"" AS ""FuelType"",
					""@BDOSFLTP"".""U_UomCode"" AS ""UomCode"",
					""@BDOSFLTP"".""U_PerKm"" AS ""PerKm"",
					""@BDOSFLTP"".""U_PerHr"" AS ""PerHr""
				FROM 
					""OITM""
				INNER JOIN 
					""OACS"" ON ""OITM"".""AssetClass"" = ""OACS"".""Code"" AND ""OACS"".""U_visCode"" = 'Y'
				INNER JOIN 
					""@BDOSFLTP"" ON ""@BDOSFLTP"".""Code"" = ""OACS"".""U_Code""
				WHERE 
					""OACS"".""Code"" = '" + itemValue + @"'");

			string query = queryBuilder.ToString();

		    oRecordSet.DoQuery(query);

			string AssetValue = Convert.ToString(oRecordSet.Fields.Item("AssetConsType").Value);
			string codeValue = oForm.Items.Item("FltCodeE").Specific.Value;
			bool emptyCode = string.IsNullOrEmpty(oForm.Items.Item("FltCodeE").Specific.Value);

			if (emptyCode || (AssetValue != codeValue))
			{			
				if (!oRecordSet.EoF)
				{
					try
					{
						SAPbouiCOM.EditText oFltCd = (SAPbouiCOM.EditText)oForm.Items.Item("FltCodeE").Specific;
						oFltCd.Value = Convert.ToString(oRecordSet.Fields.Item("ConsType").Value);
					}

					catch (Exception ex)
					{
						errorText = ex.Message;
					}

					try
					{
						SAPbouiCOM.EditText oFuelTy = (SAPbouiCOM.EditText)oForm.Items.Item("FuelTypeE").Specific;
						oFuelTy.Value = Convert.ToString(oRecordSet.Fields.Item("FuelType").Value);
					}

					catch (Exception ex)
					{
						errorText = ex.Message;
					}

					try
					{
						SAPbouiCOM.EditText oItemNm = (SAPbouiCOM.EditText)oForm.Items.Item("ItemNameE").Specific;
						oItemNm.Value = Convert.ToString(oRecordSet.Fields.Item("Name").Value);
					}

					catch (Exception ex)
					{
						errorText = ex.Message;
					}

					try
					{
						SAPbouiCOM.EditText oUomCode = (SAPbouiCOM.EditText)oForm.Items.Item("UomCodeE").Specific;
						oUomCode.Value = Convert.ToString(oRecordSet.Fields.Item("UomCode").Value);
					}

					catch (Exception ex)
					{
						errorText = ex.Message;
					}

					try
					{
						SAPbouiCOM.EditText oPerKm = (SAPbouiCOM.EditText)oForm.Items.Item("PerKmE").Specific;
						oPerKm.Value = Convert.ToString(oRecordSet.Fields.Item("PerKm").Value);
					}

					catch (Exception ex)
					{
						errorText = ex.Message;
					}

					try
					{
						SAPbouiCOM.EditText oPerHr = (SAPbouiCOM.EditText)oForm.Items.Item("PerHrE").Specific;
						oPerHr.Value = Convert.ToString(oRecordSet.Fields.Item("PerHr").Value);
					}

					catch (Exception ex)
					{
						errorText = ex.Message;
					}
				}
			}
		}
		public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction, out string errorText)
		{
			errorText = null;

			try
			{
				string sCFL_ID = oCFLEvento.ChooseFromListUID;
				SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

				if (beforeAction == false)
				{
					SAPbouiCOM.DataTable oDataTable = null;
					oDataTable = oCFLEvento.SelectedObjects;

					if (oDataTable != null)
					{
						if (sCFL_ID == "Assets_CFL")
						{
							string code = Convert.ToString(oDataTable.GetValue("Code", 0));
							string fuelType = Convert.ToString(oDataTable.GetValue("U_FuelType", 0));
							string name = Convert.ToString(oDataTable.GetValue("U_Name", 0));
							string perKm = Convert.ToString(oDataTable.GetValue("U_PerKm", 0));
							string perHr = Convert.ToString(oDataTable.GetValue("U_PerHr", 0));

							try
							{
								SAPbouiCOM.EditText oFltCd = (SAPbouiCOM.EditText)oForm.Items.Item("FltCodeE").Specific;
								oFltCd.Value = code;
							}

							catch (Exception ex)
							{
								errorText = ex.Message;
							}

							try
							{
								SAPbouiCOM.EditText oFuelty = (SAPbouiCOM.EditText)oForm.Items.Item("FuelTypeE").Specific;
								oFuelty.Value = fuelType;
							}

							catch (Exception ex)
							{
								errorText = ex.Message;
							}

							try
							{
								SAPbouiCOM.EditText oItemNm = (SAPbouiCOM.EditText)oForm.Items.Item("ItemNameE").Specific;
								oItemNm.Value = name;
							}

							catch (Exception ex)
							{
								errorText = ex.Message;
							}

							try
							{
								SAPbouiCOM.EditText oPerkm = (SAPbouiCOM.EditText)oForm.Items.Item("PerKmE").Specific;
								oPerkm.Value = perKm;
							}

							catch (Exception ex)
							{
								errorText = ex.Message;
							}

							try
							{
								SAPbouiCOM.EditText oPerHr = (SAPbouiCOM.EditText)oForm.Items.Item("PerHrE").Specific;
								oPerHr.Value = perHr;
							}

							catch (Exception ex)
							{
								errorText = ex.Message;
							}

							try
							{
								SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

								string query = @"SELECT ""UomCode"", ""ItemCode"" FROM ""OUOM"" JOIN ""OITM"" on ""OUOM"".""UomEntry"" = ""OITM"".""IUoMEntry""  WHERE ""OITM"".""ItemCode"" = '" + fuelType + @"'";
								oRecordSet.DoQuery(query);
								string uomCode = oRecordSet.Fields.Item("UomCode").Value;

								SAPbouiCOM.EditText oUomCd = (SAPbouiCOM.EditText)oForm.Items.Item("UomCodeE").Specific;
								oUomCd.Value = uomCode;
							}

							catch (Exception ex)
							{
								errorText = ex.Message;
							}
						}
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
				GC.Collect();
			}
		}
		public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
		{
			BubbleEvent = true;
			string errorText = null;

			SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

			if (oForm.TypeEx == "1473000075")
			{
				if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
				{
					formDataLoad(oForm, out errorText);
					setVisibleFormItems(oForm, out errorText);				
				}
			}
		}
		public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
		{
			BubbleEvent = true;
			string errorText = null;
			if (FormUID == "NewCostCenterForm")
			{
				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
				{
					SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

					if (pVal.ItemUID == "1")
					{
						SAPbouiCOM.Form oFormFA = CurrentForm;
						SAPbouiCOM.DBDataSource DBDataSource = oFormFA.DataSources.DBDataSources.Item(0);
						string ItemCode = DBDataSource.GetValue("ItemCode", 0).ToString();

						string CostDate = oForm.Items.Item("newDate").Specific.Value;
						string CostCode = oForm.Items.Item("ItemCode").Specific.Value;
						string CostName = oForm.Items.Item("ItemName").Specific.Value;

						CreateCostCenter(oFormFA, ItemCode, CostCode, CostName, CostDate);

						oForm.Close();

						formDataLoad(oFormFA, out errorText);
					}
				}
			}
			else
			{
				if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
				{
					SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);


					if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
					{
						createFormItems(oForm, out errorText);
						formDataLoad(oForm, out errorText);
					}

					if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.ItemUID == "BDODistTXT" & pVal.BeforeAction == true)
					{
						if (oForm.DataSources.UserDataSources.Item("BDSDistCod").ValueEx != "")
						{
							return;
						}

						CurrentForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

						createNewCreationForm();
					}

					if ((pVal.ItemUID == "1470002140" || pVal.ItemUID == "1") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
					{
						setVisibleFormItems(oForm, out errorText);
						fillItems(oForm, out errorText);
					}

					if (pVal.ItemUID == "FltCodeE" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST & pVal.BeforeAction == false)
					{
						SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
						oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

						chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
					}
				}
			}
		}
		public static void createNewCreationForm()
		{
			string errorText;
			int left = 558 + 500;
			int Top = 200 + 300;

			//ფორმის აუცილებელი თვისებები
			Dictionary<string, object> formProperties = new Dictionary<string, object>();
			formProperties.Add("UniqueID", "NewCostCenterForm");
			formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
			formProperties.Add("Title", BDOSResources.getTranslate("NewCostCenterForm"));
			formProperties.Add("Left", left);
			formProperties.Add("Width", 200);
			formProperties.Add("Top", Top);
			formProperties.Add("Height", 80);
			formProperties.Add("Modality", SAPbouiCOM.BoFormModality.fm_Modal);

			SAPbouiCOM.Form oForm;
			bool newForm;
			bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

			if (formExist == true)
			{
				if (newForm == true)
				{
					//ფორმის ელემენტების თვისებები
					Dictionary<string, object> formItems = null;

					Top = 1;
					left = 6;

					formItems = new Dictionary<string, object>();
					string itemName = "newDateS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left);
					formItems.Add("Width", 100);
					formItems.Add("Top", Top);
					formItems.Add("Height", 19);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("Date"));
					formItems.Add("Enabled", true);


					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}
					left = left + 105;

					formItems = new Dictionary<string, object>();
					itemName = "newDate";
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
					formItems.Add("Length", 30);
					formItems.Add("Size", 20);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
					formItems.Add("TableName", "");
					formItems.Add("Alias", itemName);
					formItems.Add("Bound", true);
					formItems.Add("Left", left);
					formItems.Add("Width", 100);
					formItems.Add("Top", Top);
					formItems.Add("Height", 19);
					formItems.Add("UID", itemName);
					formItems.Add("ValueEx", DateTime.Now.ToString("yyyyMMdd"));

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					Top = Top + 19 + 5;
					left = 6;

					formItems = new Dictionary<string, object>();
					itemName = "ItemCodeS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left);
					formItems.Add("Width", 100);
					formItems.Add("Top", Top);
					formItems.Add("Height", 19);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("Code"));
					formItems.Add("Enabled", true);


					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}
					left = left + 105;

					formItems = new Dictionary<string, object>();
					itemName = "ItemCode";
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("Length", 30);
					formItems.Add("Size", 20);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
					formItems.Add("TableName", "");
					formItems.Add("Alias", itemName);
					formItems.Add("Bound", true);
					formItems.Add("Left", left);
					formItems.Add("Width", 100);
					formItems.Add("Top", Top);
					formItems.Add("Height", 19);
					formItems.Add("UID", itemName);


					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					Top = Top + 19 + 5;
					left = 6;

					formItems = new Dictionary<string, object>();
					itemName = "ItemNameS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left);
					formItems.Add("Width", 100);
					formItems.Add("Top", Top);
					formItems.Add("Height", 19);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("Name"));
					formItems.Add("Enabled", true);


					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}
					left = left + 105;

					formItems = new Dictionary<string, object>();
					itemName = "ItemName";
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("Length", 30);
					formItems.Add("Size", 20);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
					formItems.Add("TableName", "");
					formItems.Add("Alias", itemName);
					formItems.Add("Bound", true);
					formItems.Add("Left", left);
					formItems.Add("Width", 100);
					formItems.Add("Top", Top);
					formItems.Add("Height", 19);
					formItems.Add("UID", itemName);


					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}
					Top = Top + 19 + 5;
					left = 6;

					itemName = "1";
					formItems = new Dictionary<string, object>();
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
					formItems.Add("Left", left);
					formItems.Add("Width", 120);
					formItems.Add("Top", Top);
					formItems.Add("Height", 19);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("Create"));

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}
				}
				oForm.ClientHeight = 100;
				oForm.ClientWidth = 250;
				oForm.Visible = true;

			}

			GC.Collect();
		}
		public static void formDataLoad(SAPbouiCOM.Form oForm, out string errorText)
		{
			errorText = null;
			string caption = BDOSResources.getTranslate("CreateDistributionRule");
			try
			{
				string Code = oForm.DataSources.DBDataSources.Item(0).GetValue("ItemCode", 0);

				//-------------------------------------------Distribution rule----------------------------------->
				string DistrRuleCode = "";
				string DistrRuleName = "";
				if (Code != "")
				{
					string Query = @"Select * from OOCR
                    inner join OPRC on OPRC.""PrcCode"" = OOCR.""OcrCode"" and OPRC.""U_BDOSFACode"" = '" + Code + "'";
					SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

					oRecordSet.DoQuery(Query);

					if (!oRecordSet.EoF)
					{
						DistrRuleCode = oRecordSet.Fields.Item("OcrCode").Value.ToString();
						DistrRuleName = oRecordSet.Fields.Item("OcrName").Value.ToString();
					}
					else
					{
						DistrRuleCode = "";
					}


					if (DistrRuleCode != "")
					{
						caption = BDOSResources.getTranslate("DistributionRule") + ": " + DistrRuleCode + "(" + DistrRuleName + ")";
					}
				}
				else
				{
					caption = BDOSResources.getTranslate("CreateDistributionRule");
					DistrRuleCode = "";
				}

				oForm.Items.Item("BDODistTXT").Specific.Caption = caption;
				oForm.DataSources.UserDataSources.Item("BDSDistCod").ValueEx = DistrRuleCode;

				//<-------------------------------------------სასაქონლო ზედნადები-----------------------------------
			}
			catch (Exception ex)
			{
				//oForm.Items.Item("BDODistTXT").Specific.Caption = caption;
				oForm.DataSources.UserDataSources.Item("BDSDistCod").ValueEx = "";
			}

		}
		public static string getFADimension(string ItemCode)
		{
			string ProfitCode = "";

			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			string Query = @"select ""OPRC"".""PrcCode"" from ""OPRC""
            inner join ""OOCR"" on ""OPRC"".""PrcCode"" = ""OOCR"".""OcrCode"" and ""OPRC"".""U_BDOSFACode"" = '" + ItemCode + "'";


			oRecordSet.DoQuery(Query);
			if (!oRecordSet.EoF)
			{
				return oRecordSet.Fields.Item("PrcCode").Value.ToString();
			}

			return ProfitCode;
		}
		public static string getAssetClassDetermination(string AssetClass)
		{
			string ClassDet = "";

			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			string Query = @"select ""ACS1"".""AcctDtn"" from ""ACS1""
            Where ""Code"" = '" + AssetClass + "'";

			oRecordSet.DoQuery(Query);
			if (!oRecordSet.EoF)
			{
				return oRecordSet.Fields.Item("AcctDtn").Value.ToString();
			}

			return ClassDet;
		}
		private static void CreateCostCenter(SAPbouiCOM.Form oForm, string ItemCode, string CostCode, string CostName, string CostDate)
		{



			SAPbobsCOM.CompanyService oCmpSrv;
			SAPbobsCOM.IProfitCentersService oProfitCentersService;
			SAPbobsCOM.IProfitCenter oProfitCenter;
			oCmpSrv = Program.oCompany.GetCompanyService();
			oProfitCentersService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService);

			string Query = @"select* from ""OPRC""
                where ""U_BDOSFACode"" = '" + ItemCode + "'";

			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			oRecordSet.DoQuery(Query);

			if (!oRecordSet.EoF)
			{
				return;
			}

			try
			{
				oProfitCenter = oProfitCentersService.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenter);

				oProfitCenter.UserFields.Item("U_BDOSFACode").Value = ItemCode;
				oProfitCenter.CenterCode = CostCode;
				oProfitCenter.CenterName = CostName;
				oProfitCenter.Effectivefrom = Convert.ToDateTime(DateTime.ParseExact(CostDate, "yyyyMMdd", CultureInfo.InvariantCulture));
				oProfitCenter.InWhichDimension = Convert.ToInt32(CommonFunctions.getOADM("U_BDOSFADim"));

				oProfitCentersService.AddProfitCenter((SAPbobsCOM.ProfitCenter)oProfitCenter);
				Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("CostCenterCreatedSuccessfully"), SAPbouiCOM.BoMessageTime.bmt_Short, false);
			}
			catch (Exception ex)
			{
				string ErrorText = ex.Message;
				Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("CostCenterNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + ErrorText);
			}
		}
	}
}
