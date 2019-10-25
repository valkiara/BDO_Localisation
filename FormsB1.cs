using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class FormsB1
    {      
        public static void allUserFieldsForAddOn(  out string errorText)
        {           
            BusinessPartners.createUserFields( out errorText);

            FixedAsset.createUserFields( out errorText);

            Items.createUserFields( out errorText);

            ItemGroup.createUserFields(out errorText);

            BDO_BPCatalog.createUserFields( out errorText);           

            Users.createUserFields( out errorText);

            BDO_RSUoM.createUserFields( out errorText);

            ARCreditNote.createUserFields( out errorText);

            BlanketAgreement.createUserFields( out errorText);

            ARInvoice.createUserFields(out errorText);

            Capitalization.createUserFields(out errorText);

            Delivery.createUserFields( out errorText);

            APCreditMemo.createUserFields( out errorText);

            VatGroup.createUserFields( out errorText);

            WithholdingTax.createUserFields( out errorText);

            LandedCosts.createUserFields( out errorText);

            HouseBankAccounts.createUserFields( out errorText);

            OutgoingPayment.createUserFields( out errorText);

            GoodsIssue.createUserFields( out errorText);

            APInvoice.createUserFields( out errorText);

            StockTransfer.createUserFields( out errorText);

            StockTransferRequest.createUserFields(out errorText);

            GoodsReceiptPO.createUserFields( out errorText);

            APDownPayment.createUserFields( out errorText);

            Retirement.createUserFields( out errorText);

            IncomingPayment.createUserFields( out errorText);

            ChartOfAccounts.createUserFields( out errorText);

            JournalEntry.createUserFields( out errorText);

            ARDownPaymentRequest.createUserFields( out errorText);
            
            GeneralSettings.createUserFields(out errorText);

            DocumentSettings.createUserFields(out errorText);

            Locations.createUserFields(out errorText);

            Warehouses.createUserFields(out errorText);

            BPBankAccounts.createUserFields(out errorText);

            AssetClass.createUserFields(out errorText);

            Projects.createUserFields(out errorText);
        }

        public static void addMenusForAddOn( out string errorText)
        {
            BDO_ProfitTaxBase.addMenus( out errorText);

            BDO_ProfitTaxBaseType.addMenus( out errorText); 
            
            BDO_WaybillsJournalSent.addMenus( out errorText);

            BDO_WaybillsJournalReceived.addMenus( out errorText);

            BDOSDepreciationAccrualWizard.addMenus(out errorText);

            BDOSFuelWriteOffWizard.addMenus(out errorText);

            BDOSTaxJournal.addMenus( out errorText);

            BDO_Drivers.addMenus( out errorText);

            BDO_Vehicles.addMenus( out errorText);

            BDOSFuelTypes.addMenus();

            BDOSFuelCriteria.addMenus();

            BDOSFuelNormSpecification.addMenus();

            BDOSFuelConsumptionAct.addMenus();

            BDO_Waybills.addMenus( out errorText);

            BDO_TaxInvoiceSent.addMenus( out errorText);

            BDOSARDownPaymentVATAccrual.addMenus( out errorText);

            BDO_ProfitTaxAccrual.addMenus( out errorText);

            BDO_TaxInvoiceReceived.addMenus( out errorText);

            BDOSInternetBanking.addMenus( out errorText);

            BDOSWaybillsAnalysisSent.addMenus( out errorText);

            BDOSWaybillsAnalysisReceived.addMenus( out errorText);

            BDOSDeleteUDF.addMenus( out errorText);

            BDOSInternetBankingIntegrationServicesRules.addMenus( out errorText);

            BDOSItemCategories.addMenus( out errorText);

            BDOSOutgoingPaymentsWizard.addMenus( out errorText);

            BDOSVATAccrualWizard.addMenus( out errorText);

            BDOSVATReconcilationWizard.addMenus( out errorText);

            BDOSFixedAssetTransfer.addMenus(out errorText);
            
            BDOSStockTransferWizard.addMenus(out errorText);

            BDOSDepreciationAccrualDocument.addMenus(out errorText);

            BDOSFuelTransferWizard.addMenus(out errorText);

            BDOSFuelConsumption.addMenus(out errorText);
        }

        public static int getLongIntRGB(int R, int G, int B)
        {
            int intValue = B * 65536 + G * 256 + R;
            return intValue;
        }

        public static bool createForm( Dictionary<string, object> formProperties, out SAPbouiCOM.Form oForm, out bool newForm, out string errorText)
        {
            errorText = null;
            object propertyValue = null;
            oForm = null;
            newForm = false;

            if (formProperties.TryGetValue("UniqueID", out propertyValue) == true)
            {
                if (formExistByIndex( propertyValue, out oForm) == true)
                {
                    return true;
                }
            }
            else
            {
                return false;
            }

            SAPbouiCOM.FormCreationParams oCreationParams = (SAPbouiCOM.FormCreationParams)(Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams));

            if (formProperties.TryGetValue("BorderStyle", out propertyValue) == true)
            {
                oCreationParams.BorderStyle = (SAPbouiCOM.BoFormBorderStyle)propertyValue;
            }

            if (formProperties.TryGetValue("FormType", out propertyValue) == true)
            {
                oCreationParams.FormType = propertyValue.ToString();
            }

            if (formProperties.TryGetValue("Modality", out propertyValue) == true)
            {
                oCreationParams.Modality = (SAPbouiCOM.BoFormModality)propertyValue;
            }

            if (formProperties.TryGetValue("ObjectType", out propertyValue) == true)
            {
                oCreationParams.ObjectType = propertyValue.ToString();
            }

            if (formProperties.TryGetValue("UniqueID", out propertyValue) == true)
            {
                oCreationParams.UniqueID = propertyValue.ToString();
            }

            if (formProperties.TryGetValue("XmlData", out propertyValue) == true)
            {
                oCreationParams.XmlData = propertyValue.ToString();
            }

            try
            {
                oForm = Program.uiApp.Forms.AddEx(oCreationParams);
                newForm = true;
            }

            catch (Exception ex)
            {
                errorText = BDOSResources.getTranslate("ErrorOfFormAdd") + " " + BDOSResources.getTranslate("Form") + " : " + "\"" + oCreationParams.UniqueID + "\"! ERROR : " + ex.Message;
                return false;
            }

            if (formProperties.TryGetValue("BrowseBy", out propertyValue) == true)
            {
                oForm.DataBrowser.BrowseBy = propertyValue.ToString();
            }

            if (formProperties.TryGetValue("DefButton", out propertyValue) == true)
            {
                oForm.DefButton = propertyValue.ToString();
            }

            if (formProperties.TryGetValue("ActiveItem", out propertyValue) == true)
            {
                oForm.ActiveItem = propertyValue.ToString();
            }

            if (formProperties.TryGetValue("AutoManaged", out propertyValue) == true)
            {
                oForm.AutoManaged = Convert.ToBoolean(propertyValue);
            }

            if (formProperties.TryGetValue("ClientHeight", out propertyValue) == true)
            {
                oForm.ClientHeight = Convert.ToInt32(propertyValue);
            }

            if (formProperties.TryGetValue("ClientWidth", out propertyValue) == true)
            {
                oForm.ClientWidth = Convert.ToInt32(propertyValue);
            }

            if (formProperties.TryGetValue("Height", out propertyValue) == true)
            {
                oForm.Height = Convert.ToInt32(propertyValue);
            }

            if (formProperties.TryGetValue("Left", out propertyValue) == true)
            {
                oForm.Left = Convert.ToInt32(propertyValue);
            }

            if (formProperties.TryGetValue("MaxHeight", out propertyValue) == true)
            {
                oForm.MaxHeight = Convert.ToInt32(propertyValue);
            }

            if (formProperties.TryGetValue("MaxWidth", out propertyValue) == true)
            {
                oForm.MaxWidth = Convert.ToInt32(propertyValue);
            }

            if (formProperties.TryGetValue("Mode", out propertyValue) == true)
            {
                oForm.Mode = (SAPbouiCOM.BoFormMode)propertyValue;
            }

            if (formProperties.TryGetValue("PaneLevel", out propertyValue) == true)
            {
                oForm.PaneLevel = Convert.ToInt32(propertyValue);
            }

            if (formProperties.TryGetValue("ReportType", out propertyValue) == true)
            {
                oForm.ReportType = propertyValue.ToString();
            }

            if (formProperties.TryGetValue("State", out propertyValue) == true)
            {
                oForm.State = (SAPbouiCOM.BoFormStateEnum)propertyValue;
            }

            if (formProperties.TryGetValue("SupportedModes", out propertyValue) == true)
            {
                oForm.SupportedModes = Convert.ToInt32(propertyValue);
            }

            if (formProperties.TryGetValue("Title", out propertyValue) == true)
            {
                oForm.Title = propertyValue.ToString();
            }

            if (formProperties.TryGetValue("Top", out propertyValue) == true)
            {
                oForm.Top = Convert.ToInt32(propertyValue);
            }

            if (formProperties.TryGetValue("Visible", out propertyValue) == true)
            {
                oForm.Visible = Convert.ToBoolean(propertyValue);
            }

            if (formProperties.TryGetValue("VisibleEx", out propertyValue) == true)
            {
                oForm.VisibleEx = Convert.ToBoolean(propertyValue);
            }

            if (formProperties.TryGetValue("Width", out propertyValue) == true)
            {
                oForm.Width = Convert.ToInt32(propertyValue);
            }

            return true;
        }

        public static void createFormItem(SAPbouiCOM.Form oForm, Dictionary<string, object> formItems, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Item oItem = null;
            object propertyValue = null;
            string UID = null;
            SAPbouiCOM.BoFormItemTypes Type = SAPbouiCOM.BoFormItemTypes.it_STATIC;

            if (formItems.TryGetValue("UID", out propertyValue) == true)
            {
                UID = propertyValue.ToString();
            }

            if (formItems.TryGetValue("Type", out propertyValue) == true)
            {
                Type = (SAPbouiCOM.BoFormItemTypes)propertyValue;
            }
            try
            {
                oItem = oForm.Items.Add(UID, Type);
            }
            catch (Exception ex)
            {
                errorText = BDOSResources.getTranslate("ErrorOfFormItemAdd") + " " + BDOSResources.getTranslate("Form") + " : " + "\"" + oForm.Title + "\", " + BDOSResources.getTranslate("FormItem") + " : " + "\"" + UID + "\"! ERROR : " + ex.Message;
                return;
            }

            bool createDataSource = false;

            if (formItems.TryGetValue("isDataSource", out propertyValue) == true)
            {
                createDataSource = Convert.ToBoolean(propertyValue);
            }

            if (createDataSource)
            {
                if (formItems.TryGetValue("DataSource", out propertyValue) == true)
                {
                    try
                    {
                        if (propertyValue.ToString() == "DBDataSources")
                        {
                            SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Add(formItems["TableName"].ToString());
                        }
                        if (propertyValue.ToString() == "DataTables")
                        {
                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Add(formItems["UniqueID"].ToString());
                        }
                        if (propertyValue.ToString() == "UserDataSources")
                        {
                            SAPbouiCOM.UserDataSource oUserDataSource = oForm.DataSources.UserDataSources.Add(UID, (SAPbouiCOM.BoDataType)formItems["DataType"], Convert.ToInt32(formItems["Length"]));
                            if (formItems.TryGetValue("Value", out propertyValue) == true)
                            {
                                oUserDataSource.Value = propertyValue.ToString();
                            }

                            if (formItems.TryGetValue("ValueEx", out propertyValue) == true)
                            {
                                oUserDataSource.ValueEx = propertyValue.ToString();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        errorText = ex.Message;
                        return;
                    }
                }
            }

            if (formItems.TryGetValue("SetAutoManaged", out propertyValue) == true)
            {
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }

            if (formItems.TryGetValue("AffectsFormMode", out propertyValue) == true)
            {
                oItem.AffectsFormMode = Convert.ToBoolean(propertyValue);
            }

            if (formItems.TryGetValue("BackColor", out propertyValue) == true)
            {
                oItem.BackColor = Convert.ToInt32(propertyValue);
            }

            if (formItems.TryGetValue("Description", out propertyValue) == true)
            {
                oItem.Description = propertyValue.ToString();
            }

            if (formItems.TryGetValue("DisplayDesc", out propertyValue) == true)
            {
                oItem.DisplayDesc = Convert.ToBoolean(propertyValue);
            }

            if (formItems.TryGetValue("Enabled", out propertyValue) == true)
            {
                oItem.Enabled = Convert.ToBoolean(propertyValue);
            }

            if (formItems.TryGetValue("FontSize", out propertyValue) == true)
            {
                oItem.FontSize = Convert.ToInt32(propertyValue);
            }

            if (formItems.TryGetValue("ForeColor", out propertyValue) == true)
            {
                oItem.ForeColor = Convert.ToInt32(propertyValue);
            }

            if (formItems.TryGetValue("FromPane", out propertyValue) == true)
            {
                oItem.FromPane = Convert.ToInt32(propertyValue);
            }

            if (formItems.TryGetValue("Height", out propertyValue) == true)
            {
                oItem.Height = Convert.ToInt32(propertyValue);
            }

            if (formItems.TryGetValue("Left", out propertyValue) == true)
            {
                oItem.Left = Convert.ToInt32(propertyValue);
            }

            if (formItems.TryGetValue("LinkTo", out propertyValue) == true)
            {
                oItem.LinkTo = propertyValue.ToString();
            }

            if (formItems.TryGetValue("RightJustified", out propertyValue) == true)
            {
                oItem.RightJustified = Convert.ToBoolean(propertyValue);
            }

            if (formItems.TryGetValue("TextStyle", out propertyValue) == true)
            {
                oItem.TextStyle = Convert.ToInt32(propertyValue);
            }

            if (formItems.TryGetValue("Top", out propertyValue) == true)
            {
                oItem.Top = Convert.ToInt32(propertyValue);
            }

            if (formItems.TryGetValue("ToPane", out propertyValue) == true)
            {
                oItem.ToPane = Convert.ToInt32(propertyValue);
            }

            if (formItems.TryGetValue("Visible", out propertyValue) == true)
            {
                oItem.Visible = Convert.ToBoolean(propertyValue);
            }

            if (formItems.TryGetValue("Width", out propertyValue) == true)
            {
                oItem.Width = Convert.ToInt32(propertyValue);
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_ACTIVE_X)
            {
                SAPbouiCOM.ActiveX oActiveX = ((SAPbouiCOM.ActiveX)(oItem.Specific));
                if (formItems.TryGetValue("ClassID", out propertyValue) == true)
                {
                    oActiveX.ClassID = propertyValue.ToString();
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            {
                SAPbouiCOM.Button oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                if (formItems.TryGetValue("Caption", out propertyValue) == true)
                {
                    oButton.Caption = propertyValue.ToString();
                }
                if (formItems.TryGetValue("Image", out propertyValue) == true)
                {
                    oButton.Type = SAPbouiCOM.BoButtonTypes.bt_Image;
                    oButton.Image = propertyValue.ToString();
                }
                if (formItems.TryGetValue("ChooseFromListUID", out propertyValue) == true)
                {
                    oButton.ChooseFromListUID = propertyValue.ToString();
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO)
            {
                SAPbouiCOM.ButtonCombo oButtonCombo = ((SAPbouiCOM.ButtonCombo)(oItem.Specific));
                if (formItems.TryGetValue("Caption", out propertyValue) == true)
                {
                    oButtonCombo.Caption = propertyValue.ToString();
                }
                if (formItems.TryGetValue("ExpandType", out propertyValue) == true)
                {
                    oButtonCombo.ExpandType = (SAPbouiCOM.BoExpandType)propertyValue;
                }
                if (formItems.TryGetValue("Bound", out propertyValue) == true)
                {
                    oButtonCombo.DataBind.SetBound(Convert.ToBoolean(formItems["Bound"]), formItems["TableName"].ToString(), formItems["Alias"].ToString());
                }
                if (formItems.TryGetValue("ValidValues", out propertyValue) == true)
                {
                    if (propertyValue.GetType() == typeof(Dictionary<string, string>))
                    {
                        Dictionary<string, string> listValidValues = (Dictionary<string, string>)propertyValue;

                        foreach (KeyValuePair<string, string> keyValue in listValidValues)
                        {
                            oButtonCombo.ValidValues.Add(keyValue.Key, keyValue.Value);
                        }
                        listValidValues = null;
                    }
                    else
                    {
                        List<string> listValidValues = (List<string>)propertyValue;

                        for (int i = 0; i < listValidValues.Count(); i++)
                        {
                            oButtonCombo.ValidValues.Add(i == 0 & listValidValues[i] == "" ? "-1" : i.ToString(), listValidValues[i]);
                        }
                        listValidValues = null;
                    }
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            {
                SAPbouiCOM.CheckBox oCheckBox = ((SAPbouiCOM.CheckBox)(oItem.Specific));
                if (formItems.TryGetValue("Caption", out propertyValue) == true)
                {
                    oCheckBox.Caption = propertyValue.ToString();
                }
                if (formItems.TryGetValue("Checked", out propertyValue) == true)
                {
                    oCheckBox.Checked = Convert.ToBoolean(propertyValue);
                }
                if (formItems.TryGetValue("ValOff", out propertyValue) == true)
                {
                    oCheckBox.ValOff = propertyValue.ToString();
                }
                if (formItems.TryGetValue("ValOn", out propertyValue) == true)
                {
                    oCheckBox.ValOn = propertyValue.ToString();
                }
                if (formItems.TryGetValue("Bound", out propertyValue) == true)
                {
                    oCheckBox.DataBind.SetBound(Convert.ToBoolean(formItems["Bound"]), formItems["TableName"].ToString(), formItems["Alias"].ToString());
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            {
                SAPbouiCOM.ComboBox oComboBox = ((SAPbouiCOM.ComboBox)(oItem.Specific));
                if (formItems.TryGetValue("Active", out propertyValue) == true)
                {
                    oComboBox.Active = Convert.ToBoolean(propertyValue);
                }
                if (formItems.TryGetValue("ExpandType", out propertyValue) == true)
                {
                    oComboBox.ExpandType = (SAPbouiCOM.BoExpandType)propertyValue;
                }
                if (formItems.TryGetValue("TabOrder", out propertyValue) == true)
                {
                    oComboBox.TabOrder = Convert.ToInt32(propertyValue);
                }
                if (formItems.TryGetValue("Bound", out propertyValue) == true)
                {
                    oComboBox.DataBind.SetBound(Convert.ToBoolean(formItems["Bound"]), formItems["TableName"].ToString(), formItems["Alias"].ToString());
                }
                if (formItems.TryGetValue("ValidValues", out propertyValue) == true)
                {
                    if (propertyValue.GetType() == typeof(Dictionary<string, string>))
                    {
                        Dictionary<string, string> listValidValues = (Dictionary<string, string>)propertyValue;

                        foreach (KeyValuePair<string, string> keyValue in listValidValues)
                        {
                            oComboBox.ValidValues.Add(keyValue.Key, keyValue.Value);
                        }
                        listValidValues = null;
                    }
                    else
                    {
                        List<string> listValidValues = (List<string>)propertyValue;

                        for (int i = 0; i < listValidValues.Count(); i++)
                        {
                            oComboBox.ValidValues.Add(i == 0 & listValidValues[i] == "" ? "-1" : i.ToString(), listValidValues[i]);
                        }
                        listValidValues = null;
                    }
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_EDIT)
            {
                SAPbouiCOM.EditText oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
                if (formItems.TryGetValue("Active", out propertyValue) == true)
                {
                    oEditText.Active = Convert.ToBoolean(propertyValue);
                }
                if (formItems.TryGetValue("BackColor", out propertyValue) == true)
                {
                    oEditText.BackColor = Convert.ToInt32(propertyValue);
                }
                if (formItems.TryGetValue("FontSize", out propertyValue) == true)
                {
                    oEditText.FontSize = Convert.ToInt32(propertyValue);
                }
                if (formItems.TryGetValue("ForeColor", out propertyValue) == true)
                {
                    oEditText.ForeColor = Convert.ToInt32(propertyValue);
                }
                if (formItems.TryGetValue("IsPassword", out propertyValue) == true)
                {
                    oEditText.IsPassword = Convert.ToBoolean(propertyValue);
                }
                if (formItems.TryGetValue("ScrollBars", out propertyValue) == true)
                {
                    oEditText.ScrollBars = (SAPbouiCOM.BoScrollBars)propertyValue;
                }
                if (formItems.TryGetValue("String", out propertyValue) == true)
                {
                    oEditText.String = propertyValue.ToString();
                }
                if (formItems.TryGetValue("SuppressZeros", out propertyValue) == true)
                {
                    oEditText.SuppressZeros = Convert.ToBoolean(propertyValue);
                }
                if (formItems.TryGetValue("TabOrder", out propertyValue) == true)
                {
                    oEditText.TabOrder = Convert.ToInt32(propertyValue);
                }
                if (formItems.TryGetValue("TextStyle", out propertyValue) == true)
                {
                    oEditText.TextStyle = Convert.ToInt32(propertyValue);
                }
                if (formItems.TryGetValue("Value", out propertyValue) == true)
                {
                    oEditText.Value = propertyValue.ToString();
                }
                if (formItems.TryGetValue("Bound", out propertyValue) == true)
                {
                    oEditText.DataBind.SetBound(Convert.ToBoolean(formItems["Bound"]), formItems["TableName"].ToString(), formItems["Alias"].ToString());
                }
                if (formItems.TryGetValue("ChooseFromListUID", out propertyValue) == true)
                {
                    oEditText.ChooseFromListUID = propertyValue.ToString();
                }
                if (formItems.TryGetValue("ChooseFromListAlias", out propertyValue) == true)
                {
                    oEditText.ChooseFromListAlias = propertyValue.ToString();
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            {
                SAPbouiCOM.EditText oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
                if (formItems.TryGetValue("Active", out propertyValue) == true)
                {
                    oEditText.Active = Convert.ToBoolean(propertyValue);
                }
                if (formItems.TryGetValue("BackColor", out propertyValue) == true)
                {
                    oEditText.BackColor = Convert.ToInt32(propertyValue);
                }
                if (formItems.TryGetValue("FontSize", out propertyValue) == true)
                {
                    oEditText.FontSize = Convert.ToInt32(propertyValue);
                }
                if (formItems.TryGetValue("ForeColor", out propertyValue) == true)
                {
                    oEditText.ForeColor = Convert.ToInt32(propertyValue);
                }
                if (formItems.TryGetValue("IsPassword", out propertyValue) == true)
                {
                    oEditText.IsPassword = Convert.ToBoolean(propertyValue);
                }
                if (formItems.TryGetValue("ScrollBars", out propertyValue) == true)
                {
                    oEditText.ScrollBars = (SAPbouiCOM.BoScrollBars)propertyValue;
                }
                if (formItems.TryGetValue("String", out propertyValue) == true)
                {
                    oEditText.String = propertyValue.ToString();
                }
                if (formItems.TryGetValue("SuppressZeros", out propertyValue) == true)
                {
                    oEditText.SuppressZeros = Convert.ToBoolean(propertyValue);
                }
                if (formItems.TryGetValue("TabOrder", out propertyValue) == true)
                {
                    oEditText.TabOrder = Convert.ToInt32(propertyValue);
                }
                if (formItems.TryGetValue("TextStyle", out propertyValue) == true)
                {
                    oEditText.TextStyle = Convert.ToInt32(propertyValue);
                }
                if (formItems.TryGetValue("Value", out propertyValue) == true)
                {
                    oEditText.Value = propertyValue.ToString();
                }
                if (formItems.TryGetValue("Bound", out propertyValue) == true)
                {
                    oEditText.DataBind.SetBound(Convert.ToBoolean(formItems["Bound"]), formItems["TableName"].ToString(), formItems["Alias"].ToString());
                }
                if (formItems.TryGetValue("ChooseFromListUID", out propertyValue) == true)
                {
                    oEditText.ChooseFromListUID = propertyValue.ToString();
                }
                if (formItems.TryGetValue("ChooseFromListAlias", out propertyValue) == true)
                {
                    oEditText.ChooseFromListAlias = propertyValue.ToString();
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            {
                SAPbouiCOM.Folder oFolder = ((SAPbouiCOM.Folder)(oItem.Specific));
                if (formItems.TryGetValue("AutoPaneSelection", out propertyValue) == true)
                {
                    oFolder.AutoPaneSelection = Convert.ToBoolean(propertyValue);
                }
                if (formItems.TryGetValue("Caption", out propertyValue) == true)
                {
                    oFolder.Caption = propertyValue.ToString();
                }
                if (formItems.TryGetValue("Pane", out propertyValue) == true)
                {
                    oFolder.Pane = Convert.ToInt32(propertyValue);
                }
                if (formItems.TryGetValue("ValOff", out propertyValue) == true)
                {
                    oFolder.ValOff = propertyValue.ToString();
                }
                if (formItems.TryGetValue("ValOn", out propertyValue) == true)
                {
                    oFolder.ValOn = propertyValue.ToString();
                }
                if (formItems.TryGetValue("GroupWith", out propertyValue) == true)
                {
                    oFolder.GroupWith(propertyValue.ToString());
                }
                if (formItems.TryGetValue("Bound", out propertyValue) == true)
                {
                    oFolder.DataBind.SetBound(Convert.ToBoolean(formItems["Bound"]), formItems["TableName"].ToString(), formItems["Alias"].ToString());
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_GRID)
            {
                SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oItem.Specific));
                if (formItems.TryGetValue("CollapseLevel", out propertyValue) == true)
                {
                    oGrid.CollapseLevel = Convert.ToInt32(propertyValue);
                }
                if (formItems.TryGetValue("DataTable", out propertyValue) == true)
                {
                    oGrid.DataTable = (SAPbouiCOM.DataTable)propertyValue;
                }
                if (formItems.TryGetValue("SelectionMode", out propertyValue) == true)
                {
                    oGrid.SelectionMode = (SAPbouiCOM.BoMatrixSelect)propertyValue;
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            {
                SAPbouiCOM.LinkedButton oLinkedButton = ((SAPbouiCOM.LinkedButton)(oItem.Specific));
                if (formItems.TryGetValue("LinkedFormXmlPath", out propertyValue) == true)
                {
                    oLinkedButton.LinkedFormXmlPath = propertyValue.ToString();
                }
                if (formItems.TryGetValue("LinkedObject", out propertyValue) == true)
                {
                    oLinkedButton.LinkedObject = (SAPbouiCOM.BoLinkedObject)propertyValue;
                }
                if (formItems.TryGetValue("LinkedObjectType", out propertyValue) == true)
                {
                    oLinkedButton.LinkedObjectType = propertyValue.ToString();
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_MATRIX)
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
                if (formItems.TryGetValue("Layout", out propertyValue) == true)
                {
                    oMatrix.Layout = (SAPbouiCOM.BoMatrixLayoutType)propertyValue;
                }
                if (formItems.TryGetValue("SelectionMode", out propertyValue) == true)
                {
                    oMatrix.SelectionMode = (SAPbouiCOM.BoMatrixSelect)propertyValue;
                }
                if (formItems.TryGetValue("TabOrder", out propertyValue) == true)
                {
                    oMatrix.TabOrder = Convert.ToInt32(propertyValue);
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON)
            {
                SAPbouiCOM.OptionBtn oOptionBtn = ((SAPbouiCOM.OptionBtn)(oItem.Specific));
                if (formItems.TryGetValue("Caption", out propertyValue) == true)
                {
                    oOptionBtn.Caption = propertyValue.ToString();
                }
                if (formItems.TryGetValue("Selected", out propertyValue) == true)
                {
                    oOptionBtn.Selected = Convert.ToBoolean(propertyValue);
                }
                if (formItems.TryGetValue("ValOff", out propertyValue) == true)
                {
                    oOptionBtn.ValOff = propertyValue.ToString();
                }
                if (formItems.TryGetValue("ValOn", out propertyValue) == true)
                {
                    oOptionBtn.ValOn = propertyValue.ToString();
                }
                if (formItems.TryGetValue("Bound", out propertyValue) == true)
                {
                    oOptionBtn.DataBind.SetBound(Convert.ToBoolean(formItems["Bound"]), formItems["TableName"].ToString(), formItems["Alias"].ToString());
                }
                if (formItems.TryGetValue("GroupWith", out propertyValue) == true)
                {
                    oOptionBtn.GroupWith(propertyValue.ToString());
                }              
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_PANE_COMBO_BOX)
            {
                SAPbouiCOM.PaneComboBox oPaneComboBox = ((SAPbouiCOM.PaneComboBox)(oItem.Specific));
                if (formItems.TryGetValue("ValOn", out propertyValue) == true)
                {
                    oPaneComboBox.Active = Convert.ToBoolean(propertyValue);
                }
                if (formItems.TryGetValue("ExpandType", out propertyValue) == true)
                {
                    oPaneComboBox.ExpandType = (SAPbouiCOM.BoExpandType)propertyValue;
                }
                if (formItems.TryGetValue("Bound", out propertyValue) == true)
                {
                    oPaneComboBox.DataBind.SetBound(Convert.ToBoolean(formItems["Bound"]), formItems["TableName"].ToString(), formItems["Alias"].ToString());
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_PICTURE)
            {
                SAPbouiCOM.PictureBox oPictureBox = ((SAPbouiCOM.PictureBox)(oItem.Specific));
                if (formItems.TryGetValue("Picture", out propertyValue) == true)
                {
                    oPictureBox.Picture = propertyValue.ToString();
                }
                if (formItems.TryGetValue("Bound", out propertyValue) == true)
                {
                    oPictureBox.DataBind.SetBound(Convert.ToBoolean(formItems["Bound"]), formItems["TableName"].ToString(), formItems["Alias"].ToString());
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_RECTANGLE)
            {

            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_STATIC)
            {
                SAPbouiCOM.StaticText oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                if (formItems.TryGetValue("Caption", out propertyValue) == true)
                {
                    oStaticText.Caption = propertyValue.ToString();
                }
            }

            if (Type == SAPbouiCOM.BoFormItemTypes.it_WEB_BROWSER)
            {
                SAPbouiCOM.WebBrowser oWebBrowser = ((SAPbouiCOM.WebBrowser)(oItem.Specific));
                if (formItems.TryGetValue("Url", out propertyValue) == true)
                {
                    oWebBrowser.Url = propertyValue.ToString();
                }
            }
        }

        public static bool formExistByUniqueID( string uniqueID, out SAPbouiCOM.Form oForm)
        {
            bool result = false;
            oForm = null;

            foreach (SAPbouiCOM.Form form in Program.uiApp.Forms)
            {
                if (form.UniqueID == uniqueID)
                {
                    oForm = form;
                    result = true;
                    break;
                }
            }
            return result;
        }

        public static bool formExistByIndex( object index, out SAPbouiCOM.Form oForm)
        {
            bool result = false;
            oForm = null;
            try
            {
                oForm = Program.uiApp.Forms.Item(index);
                result = true;
            }
            catch
            {
                result = false;
            }
            return result;
        }

        public static void addChooseFromList( SAPbouiCOM.Form oForm, bool multiSelection, string objectType, string uniqueID)
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            oCFLs = oForm.ChooseFromLists;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;

            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
            oCFLCreationParams.MultiSelection = multiSelection;
            oCFLCreationParams.ObjectType = objectType;
            oCFLCreationParams.UniqueID = uniqueID;
            oCFL = oCFLs.Add(oCFLCreationParams);
        }

        public static string ConvertDecimalToString(decimal d)
        {
            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator, NumberGroupSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberGroupSeparator};
            return d.ToString(Nfi);
        }

        public static string ConvertDecimalToStringForEditboxStrings(decimal d)
        {
            //Use this function to fill "String" property in B1 form edittexts
            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = CommonFunctions.getOADM("DecSep").ToString(), NumberGroupSeparator = CommonFunctions.getOADM("ThousSep").ToString() };
            return d.ToString(Nfi);
        }

        public static decimal StringToDecimalByGeneralSettingsSeparators(string s)
        {

            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = CommonFunctions.getOADM("DecSep").ToString(), NumberGroupSeparator = CommonFunctions.getOADM("ThousSep").ToString() };

            return Convert.ToDecimal(s, Nfi);
        }

        public static decimal cleanStringOfNonDigits(string s)
        {
            if (string.IsNullOrEmpty(s))
                return 0;

            StringBuilder sb = new StringBuilder(s.Length);
            for (int i = 0; i < s.Length; ++i)
            {
                char c = s[i];
                if ((c < '0') & (c != '.') & (c != ',') & (c != '-')) continue;
                if ((c > '9') & (c != '.') & (c != ',') & (c != '-')) continue;
                sb.Append(s[i]);
            }

            string cleaned = sb.ToString();
            bool decSepIsRp = false;
            string NewString = "";

            for (int i = cleaned.Length - 1; i >= 0; --i)
            {
                char c = cleaned[i];
                if (Char.IsNumber(c))
                {
                    NewString = c.ToString() + NewString;
                    continue;
                }
                else
                {
                    if (decSepIsRp)
                    {
                        NewString = "ThousSep" + NewString;
                    }
                    else
                    {
                        NewString = "DecSep" + NewString;
                        decSepIsRp = true;
                    }
                }
            }

            NewString = NewString.Replace("DecSep", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator);
            NewString = NewString.Replace("ThousSep", CultureInfo.InvariantCulture.NumberFormat.NumberGroupSeparator);

            try
            {
                return Convert.ToDecimal(NewString, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
        }

        public static DateTime DateFormats(string s, string dateFormat)
        {
            if (string.IsNullOrEmpty(s) == false)
            {
                DateTime dateTime = new DateTime();
                dateTime = Convert.ToDateTime(DateTime.ParseExact(s, dateFormat, CultureInfo.InvariantCulture));

                return dateTime;
            }
            else
            {
                return new DateTime();

            }
        }

        public static void SimulateRefresh()
        {
            Program.uiApp.Menus.Item("1304").Activate();
        }

        public static void resetWidthMatrixColumns(SAPbouiCOM.Form oForm, string matrixName, string firstColumnUniqueID, int wblMTRWidth)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item(matrixName).Specific));
                int visibleColumn = 0;

                foreach (SAPbouiCOM.Column oColumn in oMatrix.Columns)
                {
                    if (oColumn.Visible == true)
                    {
                        visibleColumn++;
                    }
                }

                visibleColumn = visibleColumn - 1;

                foreach (SAPbouiCOM.Column oColumn in oMatrix.Columns)
                {
                    if (oColumn.UniqueID == firstColumnUniqueID)
                    {
                        oColumn.Width = 20 - 1;
                        wblMTRWidth = wblMTRWidth - 20 - 1;
                    }
                    else
                    {
                        oColumn.Width = wblMTRWidth / visibleColumn;
                    }
                }
            }
            catch { }
        }
    }
}
