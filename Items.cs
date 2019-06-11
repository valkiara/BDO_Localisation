using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class Items
    {
        public static bool isStockItem( string ItemCode)
        {
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "SELECT \"InvntItem\",\"ItemType\" FROM \"OITM\" WHERE \"ItemCode\" = N'" + ItemCode + "'" + " AND (\"ItemType\" = 'I' OR \"ItemType\" = 'F')";
            oRecordset.DoQuery(query);

            if (!oRecordset.EoF)
            {
                return (oRecordset.Fields.Item("InvntItem").Value == "Y" || oRecordset.Fields.Item("ItemType").Value == "F");
            }
            else
            {
                return false;
            }
        }

        public static void createUserFields( out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtg1");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Level 1");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("LinkedTable", "BDOSITMCTG");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtgN1");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Name Level 1");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtg2");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Level 2");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("LinkedTable", "BDOSITMCTG");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtgN2");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Name Level 2");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtg3");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Level 3");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("LinkedTable", "BDOSITMCTG");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtgN3");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Name Level 3");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtg4");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Level 4");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("LinkedTable", "BDOSITMCTG");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtgN4");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Name Level 4");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtg5");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Level 5");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("LinkedTable", "BDOSITMCTG");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtgN5");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Name Level 5");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtg6");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Level 6");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("LinkedTable", "BDOSITMCTG");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtgN6");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Name Level 6");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtg7");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Level 7");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("LinkedTable", "BDOSITMCTG");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtgN7");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Name Level 7");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtg8");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Level 8");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("LinkedTable", "BDOSITMCTG");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtgN8");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Name Level 8");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtg9");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Level 9");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("LinkedTable", "BDOSITMCTG");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtgN9");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Name Level 9");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtg10");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Level 10");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("LinkedTable", "BDOSITMCTG");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSCtgN10");
            fieldskeysMap.Add("TableName", "OITM");
            fieldskeysMap.Add("Description", "Item Category Name Level 10");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            GC.Collect();

        }

        public static void createFormItems( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            Dictionary<string, object> formItems;
            string itemName = "";

            int pane = 20;

            //Categories (ჩანართი)
            SAPbouiCOM.Item oFolder = oForm.Items.Item("11");
            formItems = new Dictionary<string, object>();
            itemName = "BDOSCat";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            formItems.Add("Left", oFolder.Left + oFolder.Width);
            formItems.Add("Width", oFolder.Width);
            formItems.Add("Top", oFolder.Top);
            formItems.Add("Height", oFolder.Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Categories"));
            formItems.Add("Pane", pane);
            formItems.Add("ValOn", "0");
            formItems.Add("ValOff", itemName);
            formItems.Add("GroupWith", "11");
            formItems.Add("AffectsFormMode", false);
            formItems.Add("Description", BDOSResources.getTranslate("Categories"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            SAPbouiCOM.Item oItem = oForm.Items.Item("76");
            int top = oItem.Top;
            int height = oItem.Height;
            int left_s = oItem.Left;
            int width_s = oItem.Width;

            oItem = oForm.Items.Item("80");
            int left_e = oItem.Left;
            int width_e = oItem.Width;

            bool multiSelection = false;
            string objectType = "UDO_F_BDOSITMCTG_D";
            string uniqueID_CFL;

            for (int i = 1; i <= 10; i++)
            {
                top = top + height + 5;

                //Item Category Level 1
                formItems = new Dictionary<string, object>();
                itemName = "BDOSCtgS" + i; //10 characters
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                formItems.Add("Left", left_s);
                formItems.Add("Width", width_e);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("Caption", BDOSResources.getTranslate("ItemCategory") + " " + BDOSResources.getTranslate("Level") + " " + i);
                formItems.Add("FromPane", pane);
                formItems.Add("ToPane", pane);
                formItems.Add("LinkTo", "BDOSCtg" + i);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }

                uniqueID_CFL = "CFL_ItmCtg" + i;
                FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_CFL);

                ////პირობის დადება კატეგორიის არჩევის სიაზე
                //SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_CFL);
                //SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                //SAPbouiCOM.Condition oCon = oCons.Add();
                //oCon.Alias = "U_Level";
                //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //oCon.CondVal = i.ToString(); 
                //oCFL.SetConditions(oCons);

                formItems = new Dictionary<string, object>();
                itemName = "BDOSCtg" + i; //10 characters
                formItems.Add("isDataSource", true);
                formItems.Add("DataSource", "DBDataSources");
                formItems.Add("TableName", "OITM");
                formItems.Add("Alias", "U_BDOSCtg" + i);
                formItems.Add("Bound", true);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                formItems.Add("Left", left_e);
                formItems.Add("Width", 30);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("DisplayDesc", true);
                formItems.Add("FromPane", pane);
                formItems.Add("ToPane", pane);
                formItems.Add("ChooseFromListUID", uniqueID_CFL);
                formItems.Add("ChooseFromListAlias", "Code");

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }

                formItems = new Dictionary<string, object>();
                itemName = "BDOSCtgN" + i; //10 characters
                formItems.Add("isDataSource", true);
                formItems.Add("DataSource", "DBDataSources");
                formItems.Add("TableName", "OITM");
                formItems.Add("Alias", "U_BDOSCtgN" + i);
                formItems.Add("Bound", true);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                formItems.Add("Left", left_e + 32);
                formItems.Add("Width", width_e - 32);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("DisplayDesc", true);
                formItems.Add("FromPane", pane);
                formItems.Add("ToPane", pane);
                formItems.Add("Enabled", false);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }

                //golden errow
                formItems = new Dictionary<string, object>();
                itemName = "BDOSCtgL" + i; //10 characters
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                formItems.Add("Left", left_e - 20);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("FromPane", pane);
                formItems.Add("ToPane", pane);
                formItems.Add("LinkTo", "BDOSCtg" + i);
                formItems.Add("LinkedObjectType", objectType);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }
            }

            GC.Collect();
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems( oForm, out errorText);
                }

                if (pVal.ItemUID == "BDOSCat" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == true)
                {
                    oForm.PaneLevel = 20;
                    setVisibleFormItems( oForm, out errorText);
                }

                if (pVal.ItemUID != "" && pVal.ItemUID.Length > 7 && pVal.ItemUID.Substring(0, 7) == "BDOSCtg" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    int curlevel = Convert.ToInt32(pVal.ItemUID.Substring(7));
                    fillCategories( oForm, curlevel, "", "", out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.ItemUID != "" && pVal.ItemUID.Length > 7 && pVal.ItemUID.Substring(0, 7) == "BDOSCtg" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    oForm.Freeze(true);
                    SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));
                    if (pVal.BeforeAction == false)
                    {
                        int curlevel = 0;
                        string eFatherID = "";
                        string eFatherN = "";
                        chooseFromList( oForm, oCFLEvento, out curlevel, out eFatherID, out eFatherN, out errorText);

                        fillCategories( oForm, curlevel, eFatherID, eFatherN, out errorText);
                    }
                    else
                    {
                        chooseFromListBeforeAction( oForm, oCFLEvento, out errorText);
                    }
                    oForm.Freeze(false);
                }
            }
        }

        public static void chooseFromList(  SAPbouiCOM.Form oForm, SAPbouiCOM.ChooseFromListEvent oCFL,out int curlevel, out string eFatherID, out string eFatherN, out string errorText)
        {
            errorText = null;
            curlevel = 0;
            eFatherID = "";
            eFatherN = "";

            if (oCFL.ChooseFromListUID.Length > 10 && oCFL.ChooseFromListUID.Substring(0, 10) == "CFL_ItmCtg")
            {
                try
                {
                    curlevel = Convert.ToInt32(oCFL.ChooseFromListUID.Substring(10));

                    SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFL.SelectedObjects;
                    string eCode = oDataTableSelectedObjects.GetValue("Code", 0);
                    string eName = oDataTableSelectedObjects.GetValue("Name", 0);
                    eFatherID = oDataTableSelectedObjects.GetValue("U_FatherID", 0);
                    eFatherN = oDataTableSelectedObjects.GetValue("U_FatherN", 0);

                    try
                    {
                        SAPbouiCOM.EditText CategoryEdit = oForm.Items.Item("BDOSCtg" + curlevel).Specific;
                        CategoryEdit.Value = eCode;
                    }
                    catch { }
                    try
                    {
                        SAPbouiCOM.EditText CategoryEdit = oForm.Items.Item("BDOSCtgN" + curlevel).Specific;
                        CategoryEdit.Value = eName;
                    }
                    catch { }
                    
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
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }

                    setVisibleFormItems( oForm, out errorText);
                    GC.Collect();
                }
            }

        }

        public static void fillCategories(  SAPbouiCOM.Form oForm, int curlevel, string eFatherID, string eFatherN, out string errorText)
        {
            errorText = null;            

            if (string.IsNullOrEmpty(oForm.Items.Item("BDOSCtg" + curlevel).Specific.Value))
            {
                oForm.Items.Item("BDOSCtgN" + curlevel).Specific.Value = "";
            }

            try
            {
                SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                int level = curlevel - 1;

                if (string.IsNullOrEmpty(eFatherID) == false)
                {
                    for (int i = level; i > 0; i--)
                    {
                        if (oForm.Items.Item("BDOSCtg" + i).Specific.Value != eFatherID)
                        {
                            try
                            {
                                SAPbouiCOM.EditText CategoryEdit = oForm.Items.Item("BDOSCtg" + i).Specific;
                                CategoryEdit.Value = eFatherID;
                            }
                            catch { }
                            try
                            {
                                SAPbouiCOM.EditText CategoryEdit = oForm.Items.Item("BDOSCtgN" + i).Specific;
                                CategoryEdit.Value = eFatherN;
                            }
                            catch { }
                        }

                        if (string.IsNullOrEmpty(eFatherID) == false)
                        {
                            string query = @"SELECT ""U_FatherID"", ""U_FatherN"" FROM ""@BDOSITMCTG"" WHERE ""Code"" = N'" + eFatherID + @"' AND ""U_Level"" = N'" + i + "'";
                            oRecordset.DoQuery(query);

                            if (!oRecordset.EoF)
                            {
                                eFatherID = oRecordset.Fields.Item("U_FatherID").Value;
                                eFatherN = oRecordset.Fields.Item("U_FatherN").Value;
                            }
                            else
                            {
                                eFatherID = "";
                                eFatherN = "";
                            }
                        }
                    }
                }

                level = curlevel + 1;
                for (int i = level; i <= 10; i++)
                {
                    try
                    {
                        SAPbouiCOM.EditText CategoryEdit = oForm.Items.Item("BDOSCtg" + i).Specific;
                        CategoryEdit.Value = "";
                    }
                    catch { }
                    try
                    {
                        SAPbouiCOM.EditText CategoryEdit = oForm.Items.Item("BDOSCtgN" + i).Specific;
                        CategoryEdit.Value = "";
                    }
                    catch { }
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
                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }

                setVisibleFormItems( oForm, out errorText);
                GC.Collect();

            }
        }

        public static void chooseFromListBeforeAction(  SAPbouiCOM.Form oForm, SAPbouiCOM.ChooseFromListEvent oCFL, out string errorText)
        {
            errorText = null;

            if (oCFL.ChooseFromListUID.Length > 10 && oCFL.ChooseFromListUID.Substring(0, 10) == "CFL_ItmCtg")
            {
                try
                {
                    int level = Convert.ToInt32(oCFL.ChooseFromListUID.Substring(10));

                    string FatherID = "";
                    if (level > 1)
                    {
                        string FatherLevel = (level - 1).ToString();
                        FatherID = oForm.Items.Item("BDOSCtg" + FatherLevel).Specific.Value;
                    }

                    string sCFL_ID = oCFL.ChooseFromListUID;
                    SAPbouiCOM.ChooseFromList oCFL_Form = oForm.ChooseFromLists.Item(sCFL_ID);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.Alias = "U_Level";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = level.ToString();
                    oCon.Relationship = (level == 1 || string.IsNullOrEmpty(FatherID)) ? SAPbouiCOM.BoConditionRelationship.cr_NONE : SAPbouiCOM.BoConditionRelationship.cr_AND;

                    if (string.IsNullOrEmpty(FatherID) == false)
                    {
                        oCon = oCons.Add();
                        oCon.Alias = "U_FatherID";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = FatherID;
                    }

                    oCFL_Form.SetConditions(oCons);
                }
                catch (Exception ex)
                {
                    errorText = ex.Message;
                }
            }
        }

        public static void uiApp_FormDataEvent(  ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "150")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
                {
                    setVisibleFormItems( oForm, out errorText);
                }
            }
        }

        public static void setVisibleFormItems( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            try
            {
                for (int i = 1; i <= 10; i++)
                {
                    oForm.Items.Item("BDOSCtgN" + i).Enabled = false;
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

        public static string creatBatchNumbers(string itemCode, int index, out string errorText)
        {

            errorText = null;

            try
            {
                SAPbobsCOM.Recordset oRecordSetPr = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string queryPr = @"SELECT ""U_BDOSBTCNPR"" FROM ""OADM""";

                oRecordSetPr.DoQuery(queryPr);

                string BatchNumberPrefix = "";
                if (!oRecordSetPr.EoF)
                {
                    BatchNumberPrefix = oRecordSetPr.Fields.Item("U_BDOSBTCNPR").Value;
                }

                if (BatchNumberPrefix == "" || BatchNumberPrefix.Length == 1)
                {
                    errorText = BDOSResources.getTranslate("PrefixError");
                    return errorText;
                }

                int Position = BatchNumberPrefix.Length + 1;

                string BatchNumberFinal = "";


                SAPbobsCOM.Recordset oRecordSetSu = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string querySu = @"SELECT MAX(CAST(SUBSTRING(""DistNumber"", '" + Position + @"' , LENGTH(""DistNumber"")) AS INT)) as ""BatchNumber"" FROM ""OBTN"" WHERE ""DistNumber"" LIKE '" + BatchNumberPrefix + "%'";

                oRecordSetSu.DoQuery(querySu);

                if (!oRecordSetSu.EoF)
                {
                    if (oRecordSetSu.Fields.Item("BatchNumber").Value == 0)
                    {
                        BatchNumberFinal = String.Concat(BatchNumberPrefix, 10000000 + index+1);
                    }

                    else
                    {
                        BatchNumberFinal = String.Concat(BatchNumberPrefix, oRecordSetSu.Fields.Item("BatchNumber").Value + index + 1);
                    }
                }
                return BatchNumberFinal;
            }
            catch (Exception ex)
            {

                errorText = ex.ToString();
                return errorText;
            }



        }

    }
}
