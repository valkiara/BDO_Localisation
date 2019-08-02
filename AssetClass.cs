using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
	class AssetClass
	{
		public static void createUserFields(out string errorText)
		{
			errorText = null;
			Dictionary<string, object> fieldskeysMap;

			//Checkbox
			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "visCode");
			fieldskeysMap.Add("TableName", "OACS");
			fieldskeysMap.Add("Description", "Visible Code");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 1);
			fieldskeysMap.Add("DefaultValue", "N");

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			fieldskeysMap = new Dictionary<string, object>();
			fieldskeysMap.Add("Name", "Code");
			fieldskeysMap.Add("TableName", "OACS");
			fieldskeysMap.Add("Description", "Code");
			fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
			fieldskeysMap.Add("EditSize", 50);

			UDO.addUserTableFields(fieldskeysMap, out errorText);

			GC.Collect();
		}
		public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
		{
			errorText = null;

			Dictionary<string, object> formItems = new Dictionary<string, object>();

			string itemName = "";

			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			string query = @"SELECT ""U_BDOSEnbFlM"" AS ""EnableFuelMng"" FROM ""OADM"" WHERE ""U_BDOSEnbFlM"" = 'Y'";

			oRecordSet.DoQuery(query);

			if (!oRecordSet.EoF)
			{
				SAPbouiCOM.Item oItems = oForm.Items.Item("1470000017");
				SAPbouiCOM.Item oIteme = oForm.Items.Item("1470000018");

				int top = oItems.Top + (2 * oItems.Height);
				int height = oItems.Height;
				int width = oItems.Width;
				int left_s = oItems.Left;
				int left_e = oIteme.Left;
				int topCB = oItems.Top + oItems.Height;
				int widthCB = 250;

				formItems = new Dictionary<string, object>();
				itemName = "visCode";
				formItems.Add("isDataSource", true);
				formItems.Add("DataSource", "DBDataSources");
				formItems.Add("TableName", "OACS");
				formItems.Add("Alias", "U_visCode");
				formItems.Add("Bound", true);
				formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
				formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
				formItems.Add("Length", 1);
				formItems.Add("Left", left_s);
				formItems.Add("Width", widthCB);
				formItems.Add("Top", topCB);
				formItems.Add("Height", height);
				formItems.Add("UID", itemName);
				formItems.Add("Caption", BDOSResources.getTranslate("Vehicle"));
				formItems.Add("ValOff", "N");
				formItems.Add("ValOn", "Y");
				formItems.Add("DisplayDesc", true);

				FormsB1.createFormItem(oForm, formItems, out errorText);
				if (errorText != null)
				{
					return;
				}

				formItems = new Dictionary<string, object>();
				itemName = "CodeS"; //10 characters
				formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
				formItems.Add("Left", left_s);
				formItems.Add("Width", width);
				formItems.Add("Top", top);
				formItems.Add("Height", height);
				formItems.Add("UID", itemName);
				formItems.Add("Caption", BDOSResources.getTranslate("Code"));
				formItems.Add("LinkTo", "CodeE");

				FormsB1.createFormItem(oForm, formItems, out errorText);
				if (errorText != null)
				{
					return;
				}

				bool multiSelection = false;
				string objectType = "UDO_F_BDOSFLTP_T";
				string uniqueID_AssetClassCFL = "AssetClass_CFL";
				FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_AssetClassCFL);

				formItems = new Dictionary<string, object>();
				itemName = "CodeE"; //10 characters
				formItems.Add("isDataSource", true);
				formItems.Add("DataSource", "DBDataSources");
				formItems.Add("TableName", "OACS");
				formItems.Add("Alias", "U_Code");
				formItems.Add("Bound", true);
				formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
				formItems.Add("Left", left_e);
				formItems.Add("Width", width);
				formItems.Add("Top", top);
				formItems.Add("Height", height);
				formItems.Add("UID", itemName);
				formItems.Add("DisplayDesc", true);
				formItems.Add("ChooseFromListUID", uniqueID_AssetClassCFL);
				formItems.Add("ChooseFromListAlias", "Code");

				FormsB1.createFormItem(oForm, formItems, out errorText);
				if (errorText != null)
				{
					return;
				}

				formItems = new Dictionary<string, object>();
				itemName = "AssetLB"; //10 characters
				formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
				formItems.Add("Left", left_e - 20);
				formItems.Add("Top", top);
				formItems.Add("Height", height);
				formItems.Add("UID", itemName);
				formItems.Add("LinkTo", "CodeE");
				formItems.Add("LinkedObjectType", objectType);
				formItems.Add("FromPane", 0);
				formItems.Add("ToPane", 0);

				FormsB1.createFormItem(oForm, formItems, out errorText);
				if (errorText != null)
				{
					return;
				}
			}

			GC.Collect();
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
						if (sCFL_ID == "AssetClass_CFL")
						{
							string Value = Convert.ToString(oDataTable.GetValue("Code", 0));

							try
							{
								SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("CodeE").Specific;
								oEditText.Value = Value;
							}

							catch (Exception ex)
							{
								errorText = ex.Message;
							}
						}
					}

					//setVisibleFormItems(oForm, out errorText);
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
		public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
		{
			errorText = null;

			try
			{
				bool visCode = oForm.DataSources.DBDataSources.Item("OACS").GetValue("U_visCode", 0).Trim() == "Y";
				oForm.Items.Item("CodeS").Visible = visCode;
				oForm.Items.Item("CodeE").Visible = visCode;
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
		public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
		{
			BubbleEvent = true;
			string errorText = null;

			SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

			if (oForm.TypeEx == "1472000006")
			{
				if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
				{
					AssetClass.setVisibleFormItems(oForm, out errorText);
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

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
				{
					AssetClass.createFormItems(oForm, out errorText);
					AssetClass.setVisibleFormItems(oForm, out errorText);
					Program.FORM_LOAD_FOR_VISIBLE = true;
				}

				if (pVal.ItemUID == "CodeE" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST & pVal.BeforeAction == false)
				{
					SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
					oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

					AssetClass.chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
				}

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
				{
					if (pVal.ItemUID == "visCode" && pVal.BeforeAction == false)
					{
						oForm.Freeze(true);
						AssetClass.setVisibleFormItems(oForm, out errorText);
						oForm.Freeze(false);
					}
				}

				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false)
				{
					oForm.Freeze(true);
					AssetClass.setVisibleFormItems(oForm, out errorText);
					oForm.Freeze(false);
				}
			}
		}
	}
}
