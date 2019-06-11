using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class ConnectB1
    {
        public static bool connectUI(out string errorText)
        {
            errorText = null;

            try
            {
                SAPbouiCOM.SboGuiApi guiApi = new SAPbouiCOM.SboGuiApi();
                string sConnectionString;
                sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
                guiApi.Connect(sConnectionString);
                Program.uiApp = guiApi.GetApplication(-1);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return false;
            }

            return true;
        }

        public static bool connectShared(out string errorText)
        {
            errorText = null;

            try
            {
                Program.oCompany = new SAPbobsCOM.Company();
                SAPbouiCOM.SboGuiApi guiApi = new SAPbouiCOM.SboGuiApi();

                //string sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";               
                string sConnectionString = null;

                if (System.Environment.GetCommandLineArgs().Length < 2)//only one argument -> no SBO client connection string 
                {
                    sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
                }
                else//there are 2  arguments - 2nd should be connection string from SBO client ( 0 based )
                {
                    sConnectionString = Environment.GetCommandLineArgs().GetValue(1).ToString();
                }

                guiApi.Connect(sConnectionString);

                Program.uiApp = guiApi.GetApplication(-1);

                Program.oCompany = (SAPbobsCOM.Company)Program.uiApp.Company.GetDICompany();
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return false;
            }

            return true;
        }

        public static bool connectDI(string server, SAPbobsCOM.BoDataServerTypes dbServerType, string dbUserName, string dbPassword, string licenseServer, string companyDB, string userName, string password,  out string errorText)
        {
            errorText = null;

            try
            {
                Program.oCompany = new SAPbobsCOM.Company();

                Program.oCompany.Server = server;
                Program.oCompany.DbServerType = dbServerType;
                Program.oCompany.DbUserName = dbUserName;
                Program.oCompany.DbPassword = dbPassword;
                Program.oCompany.LicenseServer = licenseServer;
                Program.oCompany.CompanyDB = companyDB;
                Program.oCompany.UserName = userName;
                Program.oCompany.Password = password;

                int returnCode = Program.oCompany.Connect();

                int errCode;
                string errMsg;

                if (returnCode == 0)
                {
                    return true;
                }
                else
                {
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = BDOSResources.getTranslate("ErrorDescription")+" : " + errMsg + "! "+BDOSResources.getTranslate("Code") +" : " + errCode + "!";

                    return false;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }

            return true;

        }
    }
}
