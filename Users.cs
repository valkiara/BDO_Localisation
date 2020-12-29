using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class Users
    {
        public static void createUserFields( out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_SU");
            fieldskeysMap.Add("TableName", "OUSR");
            fieldskeysMap.Add("Description", "User Name For RS.GE");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_SP");
            fieldskeysMap.Add("TableName", "OUSR");
            fieldskeysMap.Add("Description", "Password For RS.GE");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            /////////////////
            fieldskeysMap = new Dictionary<string, object>();
            List<string> listValidValues = new List<string>();
            listValidValues.Add("No Authorization");
            listValidValues.Add("Read Only");
            listValidValues.Add("Full");

            fieldskeysMap.Add("Name", "BDOSWblAut");
            fieldskeysMap.Add("TableName", "OUSR");
            fieldskeysMap.Add("Description", "Waybill Authorization");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);
            fieldskeysMap.Add("DefaultValue", "2");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add("No Authorization");
            listValidValues.Add("Read Only");
            listValidValues.Add("Approval");

            fieldskeysMap.Add("Name", "BDOSTaxAut");
            fieldskeysMap.Add("TableName", "OUSR");
            fieldskeysMap.Add("Description", "Tax Authorization");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);
            fieldskeysMap.Add("DefaultValue", "2");
                       
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDecAtt");
            fieldskeysMap.Add("TableName", "OUSR");
            fieldskeysMap.Add("Description", "Attach Tax Invoice Declaration");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "Y");

            UDO.addUserTableFields(fieldskeysMap, out errorText);
            //////////////////////



            GC.Collect();           
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            Dictionary<string, object> formItems;
            string itemName = "";
            SAPbouiCOM.Item oItem = oForm.Items.Item("1320000001");

            int top = oItem.Top + 15;

            //სერვის მომხმარებლის სახელი (RS.GE)
            formItems = new Dictionary<string, object>();
            itemName = "SU";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 7);
            formItems.Add("Width", 178);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("Caption", BDOSResources.getTranslate("ServiceUser"));
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDO_SU");
            formItems.Add("RightJustified", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_SU";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OUSR");
            formItems.Add("Alias", "U_BDO_SU");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", 187);
            formItems.Add("Width", 163);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Description", BDOSResources.getTranslate("ServiceUser"));
            formItems.Add("RightJustified", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 15;

            //სერვის მომხმარებლის პაროლი (RS.GE) 
            formItems = new Dictionary<string, object>();
            itemName = "SP";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 7);
            formItems.Add("Width", 178);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("Caption", BDOSResources.getTranslate("ServiceUserPassword"));
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDO_SP");
            formItems.Add("RightJustified", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_SP";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OUSR");
            formItems.Add("Alias", "U_BDO_SP");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", 187);
            formItems.Add("Width", 163);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("IsPassword", true);
            formItems.Add("Description", BDOSResources.getTranslate("ServiceUserPassword"));
            formItems.Add("RightJustified", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void setVisibleFormItems(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            try
            {
                Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings( out errorText);
                if (errorText != null)
                {
                    return;
                }
                string UserType = rsSettings["UserType"];

                if (UserType == "1") //მომხმარებლის მიხედვით 
                {
                    oForm.Items.Item("SU").Visible = true;
                    oForm.Items.Item("BDO_SU").Visible = true;
                    oForm.Items.Item("SP").Visible = true;
                    oForm.Items.Item("BDO_SP").Visible = true;
                }
                else  //ორგანიზაციის მიხედვით ან ცარიელი
                {
                    oForm.Items.Item("SU").Visible = false;
                    oForm.Items.Item("BDO_SU").Visible = false;
                    oForm.Items.Item("SP").Visible = false;
                    oForm.Items.Item("BDO_SP").Visible = false;
                }
            }

            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void updateUsersRS_Info( int USERID, string BDO_SU, string BDO_SP, out string errorText)
        {
            errorText = null;
            SAPbobsCOM.Users oUSR = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);

            try
            {
                
                CommonFunctions.StartTransaction();

                //oUSR.GetByKey(USERID);
                //oUSR.UserFields.Fields.Item("U_BDO_SU").Value = BDO_SU;
                //oUSR.UserFields.Fields.Item("U_BDO_SP").Value = BDO_SP;                 

                //int returnCode = oUSR.Update();

                //int errCode;
                //string errMsg;

                //if (returnCode != 0)
                //{
                //    Program.oCompany.GetLastError(out errCode, out errMsg);
                //    errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "!";
                //}

                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string updateQuery = @"UPDATE ""OUSR""
                                            SET ""U_BDO_SU"" = N'" + BDO_SU + @"',
                                            ""U_BDO_SP"" = N'" + BDO_SP + @"'
                                        WHERE ""OUSR"".""USERID"" = N'" + USERID + "'";

                oRecordSet.DoQuery(updateQuery);               

            }

            catch
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription")+" : " + errMsg + "! "+BDOSResources.getTranslate("Code") +" : " + errCode + "!";

                CommonFunctions.EndTransaction( SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            }

            finally
            {
                CommonFunctions.EndTransaction( SAPbobsCOM.BoWfTransOpt.wf_Commit);

                Marshal.FinalReleaseComObject(oUSR);
                GC.Collect();
            }
        }
    
        public static void getUserEmployee( out string empID, out string empName, out string errorText)
        {
            errorText = null;
            empID = "";
            empName = "";

            SAPbobsCOM.Users oUSR = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = "SELECT " +
                    "\"USERID\" " +
                    "FROM \"OUSR\" " +
                    "WHERE \"USER_CODE\" = '" + Program.oCompany.UserName + "'";

                oRecordSet.DoQuery(query);
                string USERID = null;

                while (!oRecordSet.EoF)
                {
                    USERID = oRecordSet.Fields.Item("USERID").Value.ToString();
                    oRecordSet.MoveNext();
                    break;
                }

                if (USERID != null)
                {
                    query = "SELECT " +
                        "\"empID\", " +
                        "\"firstName\", " +
                        "\"lastName\"" +
                        " FROM \"OHEM\" " +
                        "WHERE \"userId\" = '" + USERID + "'";

                    oRecordSet.DoQuery(query);
                    while (!oRecordSet.EoF)
                    {
                        empID = oRecordSet.Fields.Item("empID").Value.ToString();
                        empName = oRecordSet.Fields.Item("firstName").Value.ToString() + " " + oRecordSet.Fields.Item("lastName").Value.ToString();
                        oRecordSet.MoveNext();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }           
            finally 
            {
                Marshal.FinalReleaseComObject(oUSR);
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        public static void getUserByCode( string userCode, out string userName,out int userID,out string errorText)
        {
            errorText = null;
            userName = "";
            userID = 0;

            SAPbobsCOM.Users oUSR = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = "SELECT " +
                    "\"USERID\",\"U_NAME\" " +
                    "FROM \"OUSR\" " +
                    "WHERE \"USER_CODE\" = '" + userCode + "'";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    userName = oRecordSet.Fields.Item("U_NAME").Value;
                    userID = oRecordSet.Fields.Item("USERID").Value;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oUSR);
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
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
                    Users.createFormItems(oForm, out errorText);
                    Users.setVisibleFormItems( oForm, out errorText);
                }
            }
        }
    }
}
