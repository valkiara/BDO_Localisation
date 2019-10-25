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
            fieldskeysMap.Add("Name", "BDOSFuTp");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Fuel Type Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
		}

		public static void createFormItems(SAPbouiCOM.Form oForm)
		{
            string errorText;

            Dictionary<string, object> formItems;
			string itemName;
            int left_s = oForm.Items.Item("1470002142").Left;
            int left_e = oForm.Items.Item("1470002152").Left;
            int width_s = oForm.Items.Item("1470002142").Width;
            int width_e = oForm.Items.Item("1470002152").Width;
            int height = oForm.Items.Item("1470002142").Height;
            int top = oForm.Items.Item("1470002142").Top;
            int fromPane = oForm.Items.Item("1470002142").FromPane;
            int toPane = oForm.Items.Item("1470002142").ToPane;

            FormsB1.addChooseFromList(oForm, false, "UDO_F_BDOSFUTP_D", "FuelTypeCodeCFL"); //Fuel Types

            top += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "FUTPCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("FuelType"));
            formItems.Add("FromPane", fromPane);
            formItems.Add("ToPane", toPane);
            formItems.Add("LinkTo", "FUTPCodeE");
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "FUTPCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OITM");
            formItems.Add("Alias", "U_BDOSFuTp");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", fromPane);
            formItems.Add("ToPane", toPane);
            formItems.Add("ChooseFromListUID", "FuelTypeCodeCFL");
            formItems.Add("ChooseFromListAlias", "Code");
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "FUTPCodeLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "FUTPCodeE");
            formItems.Add("LinkedObjectType", "UDO_F_BDOSFUTP_D");
            formItems.Add("FromPane", fromPane);
            formItems.Add("ToPane", toPane);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

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
                throw new Exception(errorText);
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
            //    throw new Exception(errorText);
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

            //FormsB1.createFormItem(oForm, formItems, out errorText);
            //if (errorText != null)
            //{
            //    throw new Exception(errorText);
            //}
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                string enableFuelMng = (string)CommonFunctions.getOADM("U_BDOSEnbFlM");

                if (enableFuelMng == "Y")
                {
                    oForm.Items.Item("FUTPCodeS").Visible = true;
                    oForm.Items.Item("FUTPCodeE").Visible = true;
                    oForm.Items.Item("FUTPCodeLB").Visible = true;
                }
            }

            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
            {
                setVisibleFormItems(oForm);
            }
        }

		public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
		{
			BubbleEvent = true;
			string errorText;

			if (FormUID == "NewCostCenterForm")
			{
				if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
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

					if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
					{
						createFormItems(oForm);
						formDataLoad(oForm, out errorText); //?????????
                        Program.FORM_LOAD_FOR_VISIBLE = true;
                        Program.FORM_LOAD_FOR_ACTIVATE = true;
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE && !pVal.BeforeAction)
                    {
                        if (Program.FORM_LOAD_FOR_VISIBLE)
                        {
                            Program.FORM_LOAD_FOR_VISIBLE = false;
                            setVisibleFormItems(oForm);
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "BDODistTXT" && pVal.BeforeAction)
					{
						if (oForm.DataSources.UserDataSources.Item("BDSDistCod").ValueEx != "")
						{
							return;
						}

						CurrentForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

						createNewCreationForm();
					}

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.ItemUID == "FUTPCodeE")
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                        chooseFromList(oForm, pVal, oCFLEvento);
                    }
                }
			}
		}

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction)
                {
                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (oCFLEvento.ChooseFromListUID == "FuelTypeCodeCFL")
                        {
                            string fuelTypeCode = oDataTable.GetValue("Code", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("FUTPCodeE").Specific.Value = fuelTypeCode);
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }                        
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
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
