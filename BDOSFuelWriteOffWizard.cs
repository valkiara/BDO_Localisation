using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Data;

namespace BDO_Localisation_AddOn
{
	class BDOSFuelWriteOffWizard
	{
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

			//ფორმის აუცილებელი თვისებები
			Dictionary<string, object> formProperties = new Dictionary<string, object>();
			formProperties.Add("UniqueID", "BDOSFuelWOForm");
			formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
			formProperties.Add("Title", BDOSResources.getTranslate("FuelWriteOffWizard"));
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
					itemName = "DocDateS"; //10 characters
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
					itemName = "DocDate";
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

					multiSelection = false;
					string uniqueID_ExpAcct = "ExpAcc_CFL";
					string objectTypeExp = "1";
					FormsB1.addChooseFromList(oForm, multiSelection, objectTypeExp, uniqueID_ExpAcct);

					formItems = new Dictionary<string, object>();
					itemName = "ExpAccS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s + 400);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("ExpenseAccount"));
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "ExpAcc"; //10 characters
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("TableName", "");
					formItems.Add("Length", 20);
					formItems.Add("Size", 20);
					formItems.Add("Alias", "ExpAcc");
					formItems.Add("Bound", true);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
					formItems.Add("Left", left_e + 370);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("DisplayDesc", true);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);
					formItems.Add("ChooseFromListUID", uniqueID_ExpAcct);
					formItems.Add("ChooseFromListAlias", "AcctCode");

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}
					formItems = new Dictionary<string, object>();
					itemName = "ExpAccLB"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
					formItems.Add("Left", left_e + 350);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("LinkTo", "ExpAcc");
					formItems.Add("LinkedObjectType", objectTypeExp);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

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
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("FromDate"));
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "FromDate";
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

					formItems = new Dictionary<string, object>();
					itemName = "ToDateS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("ToDate"));
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "ToDate";
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

					multiSelection = false;
					string uniqueID_lf_ConsT = "ConsT_CFL";
					string objectTypeConsT = "UDO_F_BDOSFLTP_T";
					FormsB1.addChooseFromList(oForm, multiSelection, objectTypeConsT, uniqueID_lf_ConsT);

					top = top + height + 1;

					formItems = new Dictionary<string, object>();
					itemName = "ConsTypeS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s);
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
					itemName = "ConsType"; //10 characters
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("TableName", "");
					formItems.Add("Length", 20);
					formItems.Add("Size", 20);
					formItems.Add("Alias", "ConsType");
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
					formItems.Add("ChooseFromListUID", uniqueID_lf_ConsT);
					formItems.Add("ChooseFromListAlias", "Code");

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "ConsTypeLB"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
					formItems.Add("Left", left_e - 20);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("LinkTo", "ConsType");
					formItems.Add("LinkedObjectType", objectTypeConsT);
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					string uniqueID_lf_ItemMTR_CFL = "ItemMTR_CFL";
					string objectTypeItem = "4";
					FormsB1.addChooseFromList(oForm, true, objectTypeItem, uniqueID_lf_ItemMTR_CFL);

					//პირობის დადება ძს არჩევის სიაზე

					SAPbouiCOM.ChooseFromList oCFL_Item = oForm.ChooseFromLists.Item(uniqueID_lf_ItemMTR_CFL);
					SAPbouiCOM.Conditions oCons_Item = oCFL_Item.GetConditions();
					SAPbouiCOM.Condition oCon_Item = oCons_Item.Add();
					oCon_Item.Alias = "ItemType";
					oCon_Item.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
					oCon_Item.CondVal = "F"; //Fixed Assets
					oCFL_Item.SetConditions(oCons_Item);

					top = top + height + 1;

					formItems = new Dictionary<string, object>();
					itemName = "VehicleS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s);
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
					itemName = "Vehicle"; //10 characters
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("TableName", "");
					formItems.Add("Length", 20);
					formItems.Add("Size", 20);
					formItems.Add("Alias", "Vehicle");
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
					formItems.Add("ChooseFromListUID", uniqueID_lf_ItemMTR_CFL);
					formItems.Add("ChooseFromListAlias", "ItemCode");

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "VehicleLB"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
					formItems.Add("Left", left_e - 20);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("LinkTo", "Vehicle");
					formItems.Add("LinkedObjectType", objectTypeItem);
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
					formItems.Add("Left", left_s);
					formItems.Add("Width", width);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("FuelGroup"));
					formItems.Add("LinkTo", "FuelGroup");
					formItems.Add("FromPane", 0);
					formItems.Add("ToPane", 0);

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					Dictionary<string, string> listValidValuesItemGroups = getItemGroupsList(out errorText);

					formItems = new Dictionary<string, object>();
					itemName = "FuelGroup";
					formItems.Add("Size", 20);
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
					formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
					formItems.Add("DisplayDesc", true);
					formItems.Add("Left", left_e);
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

					string uniqueID_lf_FuelTMTR_CFL = "FuelTMTR_CFL";
					string objectType = "4";
					FormsB1.addChooseFromList(oForm, true, objectType, uniqueID_lf_FuelTMTR_CFL);

					//პირობის დადება აითემის არჩევის სიაზე

					oCFL_Item = oForm.ChooseFromLists.Item(uniqueID_lf_FuelTMTR_CFL);
					oCons_Item = oCFL_Item.GetConditions();
					oCon_Item = oCons_Item.Add();
					oCon_Item.Alias = "InvntItem";
					oCon_Item.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
					oCon_Item.CondVal = "Y";
					oCFL_Item.SetConditions(oCons_Item);

					formItems = new Dictionary<string, object>();
					itemName = "FuelTypeS"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
					formItems.Add("Left", left_s);
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
					itemName = "FuelType"; //10 characters
					formItems.Add("isDataSource", true);
					formItems.Add("DataSource", "UserDataSources");
					formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
					formItems.Add("TableName", "");
					formItems.Add("Length", 20);
					formItems.Add("Size", 20);
					formItems.Add("Alias", "FuelType");
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
					formItems.Add("ChooseFromListUID", uniqueID_lf_FuelTMTR_CFL);
					formItems.Add("ChooseFromListAlias", "ItemCode");

					FormsB1.createFormItem(oForm, formItems, out errorText);
					if (errorText != null)
					{
						return;
					}

					formItems = new Dictionary<string, object>();
					itemName = "FuelTypeLB"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
					formItems.Add("Left", left_e - 20);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("LinkTo", "FuelType");
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

					//საკონტროლო პანელი
					formItems = new Dictionary<string, object>();
					itemName = "fillMTR"; //10 characters
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

					oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50);
					oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1);
					oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("AssetCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("AssetName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("ConsumType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("StartUnit", SAPbouiCOM.BoFieldsType.ft_Quantity, 50);
					oDataTable.Columns.Add("EndUnit", SAPbouiCOM.BoFieldsType.ft_Quantity, 50);
					oDataTable.Columns.Add("WorkHours", SAPbouiCOM.BoFieldsType.ft_Quantity, 50);
					oDataTable.Columns.Add("FuelType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("FuelName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("UomEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("NormConsum", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("Consum", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("ExpAcct", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("Project", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("Dimension1", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("Dimension2", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("Dimension3", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("Dimension4", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("Dimension5", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("FuelCDoc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
					oDataTable.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);

					multiSelection = false;
					string uniqueID_lf_Acct_CFL = "AcctCode_CFL";
					FormsB1.addChooseFromList(oForm, multiSelection, "1", uniqueID_lf_Acct_CFL);
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
						else if (columnName == "ExpAcct")
						{
							oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
							oColumn.TitleObject.Caption = BDOSResources.getTranslate("ExpenseAccount");
							oColumn.Editable = true;
							oColumn.DataBind.Bind("FuelMTR", columnName);
							oColumn.ChooseFromListUID = uniqueID_lf_Acct_CFL;
							oColumn.ChooseFromListAlias = "AcctCode";
							oColumn.AffectsFormMode = false;
						}
						else if (columnName == "FuelCDoc")
						{
							oColumn = oColumns.Add("FuelCDoc", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
							oColumn.TitleObject.Caption = BDOSResources.getTranslate("FuelCDoc");
							oColumn.Editable = false;
							oColumn.DataBind.Bind("FuelMTR", columnName);
							oLink = oColumn.ExtendedObject;
							oLink.LinkedObjectType = "UDO_F_BDOSFUECON_D";
						}
						else if (columnName == "DocNum")
						{
							oColumn = oColumns.Add("DocNum", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
							oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocNum");
							oColumn.Editable = false;
							oColumn.DataBind.Bind("FuelMTR", columnName);
							oLink = oColumn.ExtendedObject;
							oLink.LinkedObjectType = "60";
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
					itemName = "CreatDoc"; //10 characters
					formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
					formItems.Add("Left", left_s);
					formItems.Add("Width", 150);
					formItems.Add("Top", top);
					formItems.Add("Height", height);
					formItems.Add("UID", itemName);
					formItems.Add("Caption", BDOSResources.getTranslate("CreateGoodsIssue"));
					//formItems.Add("SetAutoManaged", true);

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
		public static int getDocNum(string fType) {
			int docNum = 0;
			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			string query = @"SELECT ""DocNum"" FROM ""OIGE"" LEFT JOIN ""IGE1"" ON ""IGE1"".""DocEntry"" = ""OIGE"".""DocEntry"" WHERE ""IGE1"".""ItemCode"" = '" + fType +@"'";
			oRecordSet.DoQuery(query);
			if (!oRecordSet.EoF) {
				docNum = oRecordSet.Fields.Item("DocNum").Value;
			}
			return docNum;

		}
		private static void CreateGoodsIssue(SAPbouiCOM.Form oForm)
		{
			int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreateGoodsIssue") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

			if (answer == 2)
			{
				return;
			}

			SAPbouiCOM.EditText oEditTextDocDate = (SAPbouiCOM.EditText)oForm.Items.Item("DocDate").Specific;
			string DocDateS = oEditTextDocDate.Value;
			DateTime DocDate = Convert.ToDateTime(DateTime.ParseExact(DocDateS, "yyyyMMdd", CultureInfo.InvariantCulture));
			SAPbobsCOM.Documents oGoodsIssue = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
			SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("FuelMTR").Specific;
			oMatrix.FlushToDataSource();
			SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("FuelMTR");
			oGoodsIssue.DocDate = DocDate;
			SAPbouiCOM.DBDataSources oDBDataSources = oForm.DataSources.DBDataSources;
			string checkBox;

			double quantity;

			for (int i = 0; i < oDataTable.Rows.Count; i++)
			{
				checkBox = oDataTable.GetValue("CheckBox", i);
				string fType = oDataTable.GetValue("FuelType", i);
				if (checkBox=="Y")
				{
					if (fType == "")
					{
						continue;
					}
					oGoodsIssue.Lines.ItemCode = fType;
					oGoodsIssue.Lines.ItemDescription = oDataTable.GetValue("AssetName", i);

					if (oDataTable.GetValue("Consum", i) != "")
					{

						quantity = Convert.ToDouble(oDataTable.GetValue("Consum", i));
						oGoodsIssue.Lines.Quantity = quantity;
					}
					oGoodsIssue.Lines.AccountCode = oDataTable.GetValue("ExpAcct", i);
					oGoodsIssue.Lines.ProjectCode = oDataTable.GetValue("Project", i);
					oGoodsIssue.Lines.CostingCode = oDataTable.GetValue("Dimension1", i);
					oGoodsIssue.Lines.CostingCode2 = oDataTable.GetValue("Dimension2", i);
					oGoodsIssue.Lines.CostingCode3 = oDataTable.GetValue("Dimension3", i);
					oGoodsIssue.Lines.CostingCode4 = oDataTable.GetValue("Dimension4", i);
					oGoodsIssue.Lines.CostingCode5 = oDataTable.GetValue("Dimension5", i);

					oGoodsIssue.Lines.Add();
				}
			}

			int resultCode = oGoodsIssue.Add();
			int docEntry = 0;
			if (resultCode != 0)
			{
				string errorMessage = "";
				Program.oCompany.GetLastError(out resultCode, out errorMessage);
				Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + errorMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
			}
			else
			{

				bool newDoc = oGoodsIssue.GetByKey(Convert.ToInt32(Program.oCompany.GetNewObjectKey()));
				if (newDoc == true)
				{
					docEntry = oGoodsIssue.DocEntry;
					Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCreatedSuccesfully") + ": " + docEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

				}
				else
				{
					Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated"));
					return;
				}

				for (int i = 0; i < oDataTable.Rows.Count; i++)
				{
					checkBox = oDataTable.GetValue("CheckBox", i);
					if ((oDataTable.GetValue("DocNum", i) == "0"|| string.IsNullOrEmpty(oDataTable.GetValue("DocNum", i))) && checkBox == "Y")
					{
						oDataTable.SetValue("DocNum", i, docEntry.ToString());
						//Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCreatedSuccesfully") + ": " + docNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success); 
					}
					//else if(oDataTable.GetValue("DocNum", i) != "0" && checkBox == "Y")
					//{
					//	Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated"));
					//	break;
					//}
				}
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
			int rows = oMatrix.RowCount;
			//oDataTable.Rows.Clear();

			SAPbouiCOM.EditText oEditTextDate = (SAPbouiCOM.EditText)oForm.Items.Item("FromDate").Specific;
			String fromDate = oEditTextDate.Value;
			if (string.IsNullOrEmpty(fromDate))
			{
				errorText = BDOSResources.getTranslate("FromDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
				Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
				return;
			}

			SAPbouiCOM.EditText oEditTextDate2 = (SAPbouiCOM.EditText)oForm.Items.Item("ToDate").Specific;
			String toDate = oEditTextDate2.Value;
			if (string.IsNullOrEmpty(toDate))
			{
				errorText = BDOSResources.getTranslate("ToDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
				Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
				return;
			}

			string oConsType;
			oConsType = oForm.Items.Item("ConsType").Specific.Value;
			bool cons = string.IsNullOrEmpty(oConsType);

			string oVehicle;
			oVehicle = oForm.Items.Item("Vehicle").Specific.Value;
			bool veh = string.IsNullOrEmpty(oVehicle);

			string oFuelGroup;
			oFuelGroup = oForm.Items.Item("FuelGroup").Specific.Value;
			bool group = string.IsNullOrEmpty(oFuelGroup);

			string oFuelType;
			oFuelType = oForm.Items.Item("FuelType").Specific.Value;
			bool type = string.IsNullOrEmpty(oFuelType);

			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			try
			{
				//SAPbouiCOM.DataTable oDataTableFilter = oForm.DataSources.DataTables.Item("FiltTable");
				DateTime from = DateTime.ParseExact(fromDate, "yyyyMMdd", null);
				DateTime to = DateTime.ParseExact(toDate, "yyyyMMdd", null);

				StringBuilder queryBuilder = new StringBuilder();
				queryBuilder.Append(
					@"select 
                            ""@BDOSFUCON1"".""U_AssetCode"" as ""AssCode"",
                       		""@BDOSFUCON1"".""U_AssetName"" as ""AssName"",
                       		""@BDOSFUCON1"".""U_ConsumType"" as ""ConType"",
                       		""@BDOSFUCON1"".""U_StartUnit"" as ""StartUnit"",
                       		""@BDOSFUCON1"".""U_EndUnit"" as ""EndUnit"",
                       		""@BDOSFUCON1"".""U_WorkHours"" as ""WorkHours"",
                       		""@BDOSFUCON1"".""U_FuelType"" as ""FuelTyp"",
                       		""@BDOSFUCON1"".""U_FuelName"" as ""FuelName"",
                       		""@BDOSFUCON1"".""U_Uom"" as ""Uom"",
                       		""@BDOSFUCON1"".""U_NormConsum"" as ""NormConsum"",
                       		""@BDOSFUCON1"".""U_Consum"" as ""Consumption"",
                       		""@BDOSFUCON1"".""U_Project"" as ""Project"",
                       		""@BDOSFUCON1"".""U_Dimension1"" as ""Dimension1"",
                       		""@BDOSFUCON1"".""U_Dimension2"" as ""Dimension2"",
                       		""@BDOSFUCON1"".""U_Dimension3"" as ""Dimension3"",
                       		""@BDOSFUCON1"".""U_Dimension4"" as ""Dimension4"",
                       		""@BDOSFUCON1"".""U_Dimension5"" as ""Dimension5"",
                            ""@BDOSFUECON"".""DocEntry"" AS ""Entry"",
							""@BDOSFUCON1"".""U_DocNum"" AS ""DocNum"",
							""@BDOSFUECON"".""U_DocDate"" as ""U_DocDate"",
                              ""OITM"".""ItmsGrpCod"" as ""FuelGroup""
                    FROM 
                        ""@BDOSFUCON1""
                    INNER JOIN 
                        ""@BDOSFUECON"" ON ""@BDOSFUECON"".""DocEntry"" = ""@BDOSFUCON1"".""DocEntry"" 
                    INNER JOIN 
                        ""OITM"" ON ""@BDOSFUCON1"".""U_AssetCode"" = ""OITM"".""ItemCode""
                    WHERE 
                        ""@BDOSFUECON"".""U_DocDate""<='");


				queryBuilder.Append(to.ToString("yyyyMMdd"));
				queryBuilder.Append("' ");
				queryBuilder.Append(@" AND ""@BDOSFUECON"".""U_DocDate"" >= '");
				queryBuilder.Append(from.ToString("yyyyMMdd"));
				queryBuilder.Append("' ");

				if (!cons)
				{
					queryBuilder.Append(@" AND ""@BDOSFUCON1"".""U_ConsumType"" = '" + oConsType + @"'");
				}

				if (!veh)
				{
					queryBuilder.Append(@" AND ""@BDOSFUCON1"".""U_AssetCode"" = '" + oVehicle + @"'");
				}

				if (!group)
				{
					queryBuilder.Append(@" AND ""OITM"".""ItmsGrpCod"" = '" + oFuelGroup + @"'");
				}

				if (!type)
				{
					queryBuilder.Append(@" AND ""@BDOSFUCON1"".""U_FuelType"" = '" + oFuelType + @"'");
				}


				oRecordSet.DoQuery(queryBuilder.ToString());

				int count = oRecordSet.RecordCount;
				string expAcc = "";

				int rowIndex = 0;

				while (!oRecordSet.EoF)
				{
					oDataTable.Rows.Add();
					oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
					oDataTable.SetValue("CheckBox", rowIndex, "N");
					oDataTable.SetValue("DocEntry", rowIndex, oRecordSet.Fields.Item("Entry").Value);
					oDataTable.SetValue("AssetCode", rowIndex, oRecordSet.Fields.Item("AssCode").Value);
					oDataTable.SetValue("AssetName", rowIndex, oRecordSet.Fields.Item("AssName").Value);
					oDataTable.SetValue("ConsumType", rowIndex, oRecordSet.Fields.Item("ConType").Value);
					oDataTable.SetValue("StartUnit", rowIndex, oRecordSet.Fields.Item("StartUnit").Value);
					oDataTable.SetValue("EndUnit", rowIndex, oRecordSet.Fields.Item("EndUnit").Value);
					oDataTable.SetValue("WorkHours", rowIndex, oRecordSet.Fields.Item("WorkHours").Value);
					oDataTable.SetValue("FuelType", rowIndex, oRecordSet.Fields.Item("FuelTyp").Value);
					oDataTable.SetValue("FuelName", rowIndex, oRecordSet.Fields.Item("FuelName"));
					oDataTable.SetValue("UomEntry", rowIndex, oRecordSet.Fields.Item("Uom").Value);
					oDataTable.SetValue("NormConsum", rowIndex, oRecordSet.Fields.Item("NormConsum").Value);
					oDataTable.SetValue("Consum", rowIndex, oRecordSet.Fields.Item("Consumption").Value);

					expAcc = oForm.Items.Item("ExpAcc").Specific.Value;
					if (expAcc != null)
					{
						oDataTable.SetValue("ExpAcct", rowIndex, expAcc);
					}

					oDataTable.SetValue("Project", rowIndex, oRecordSet.Fields.Item("Project").Value);
					oDataTable.SetValue("Dimension1", rowIndex, oRecordSet.Fields.Item("Dimension1").Value);
					oDataTable.SetValue("Dimension2", rowIndex, oRecordSet.Fields.Item("Dimension2").Value);
					oDataTable.SetValue("Dimension3", rowIndex, oRecordSet.Fields.Item("Dimension3").Value);
					oDataTable.SetValue("Dimension4", rowIndex, oRecordSet.Fields.Item("Dimension4").Value);
					oDataTable.SetValue("Dimension5", rowIndex, oRecordSet.Fields.Item("Dimension5").Value);
					oDataTable.SetValue("FuelCDoc", rowIndex, oRecordSet.Fields.Item("Entry").Value);
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

					chooseFromList(oForm, pVal.BeforeAction, oCFLEvento, out errorText);

				}
				if ((pVal.ItemUID == "Check" || pVal.ItemUID == "Uncheck") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
				{
					checkUncheckTables(oForm, pVal.ItemUID, "FuelMTR", out errorText);
				}

				if (pVal.ItemUID == "fillMTR" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
				{
					fillMTRItems(oForm, out errorText);
				}
				if (pVal.ItemUID == "CreatDoc" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
				{
					CreateGoodsIssue(oForm);
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
		private static void chooseFromList(SAPbouiCOM.Form oForm, bool BeforeAction, SAPbouiCOM.IChooseFromListEvent oCFLEvento, out string errorText)
		{
			errorText = null;

			string sCFL_ID = oCFLEvento.ChooseFromListUID;
			SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
			SAPbouiCOM.DataTable oDataTable = null;
			oDataTable = oCFLEvento.SelectedObjects;


			if (BeforeAction == false)
			{
				//SAPbouiCOM.DataTable oDataTable = null;
				oDataTable = oCFLEvento.SelectedObjects;

				if (oDataTable != null)
				{
					try
					{
						if (sCFL_ID == "ConsT_CFL")
						{
							string code = Convert.ToString(oDataTable.GetValue("Code", 0));

							try
							{
								SAPbouiCOM.EditText oConsType = oForm.Items.Item("ConsType").Specific;
								oConsType.Value = code;
							}
							catch { }
						}

						if (sCFL_ID == "ItemMTR_CFL")
						{
							string code = Convert.ToString(oDataTable.GetValue("ItemCode", 0));

							try
							{
								SAPbouiCOM.EditText oVehicle = oForm.Items.Item("Vehicle").Specific;
								oVehicle.Value = code;
							}
							catch { }
						}

						if (sCFL_ID == "FuelTMTR_CFL")
						{
							string code = Convert.ToString(oDataTable.GetValue("ItemCode", 0));

							try
							{
								SAPbouiCOM.EditText oFuelTp = oForm.Items.Item("FuelType").Specific;
								oFuelTp.Value = code;
							}
							catch { }
						}

						if (sCFL_ID == "ExpAcc_CFL")
						{
							string code = Convert.ToString(oDataTable.GetValue("AcctCode", 0));

							try
							{
								SAPbouiCOM.EditText oExpAcct = oForm.Items.Item("ExpAcc").Specific;
								oExpAcct.Value = code;
							}
							catch { }
						}

						if (oCFLEvento.ChooseFromListUID == "AcctCode_CFL")
						{
							string acctCode = Convert.ToString(oDataTable.GetValue("AcctCode", 0));
							SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("FuelMTR").Specific));
							SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
							if (cellPos == null)
							{
								return;
							}

							SAPbouiCOM.EditText oEditText;

							try
							{
								oEditText = oMatrix.Columns.Item("ExpAcct").Cells.Item(cellPos.rowIndex).Specific;
								oEditText.Value = acctCode;
							}
							catch { }
							oMatrix.FlushToDataSource();
						}
					}
					catch (Exception ex)
					{
						//    setWhtCodes(oForm);
						//    fillMTRInvoice(oForm);
					}

				}
			}
			else
			{
				if (sCFL_ID == "ItemMTR_CFL")
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
					oCreationPackage.UniqueID = "BDOSFuelWOForm";
					oCreationPackage.String = BDOSResources.getTranslate("FuelWriteOffWizard");
					oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

					menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
				}

				catch (Exception ex)
				{
					errorText = ex.Message;
				}
			}
		}
	}
}
