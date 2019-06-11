using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn.TBC_Integration_Services
{
    static partial class MainChangePasswordService
    {
        /// <summary>პაროლის შეცვლა</summary>
        /// <param name="oChangePasswordService"></param>
        /// <param name="newPassword"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public static string changePassword(ChangePasswordService oChangePasswordService, string newPassword, out string errorText)
        {
            errorText = null;
            string changePassword = null;

            try
            {
                changePassword = oChangePasswordService.ChangePassword(newPassword);
            }
            catch (Exception ex)
            {
                try
                {
                    errorText = ex.Message + '\n' + ((System.Web.Services.Protocols.SoapException)ex).Code.Name;
                    return changePassword;
                }
                catch
                {
                    errorText = ex.Message;
                    return changePassword;
                }
            }
            return changePassword;
        }

       /// <summary>ავტორიზაციის პარამეტრების შევსება</summary>
       /// <param name="serviceUrl"></param>
       /// <param name="username"></param>
       /// <param name="password"></param>
       /// <param name="nonce"></param>
       /// <returns></returns>
        public static ChangePasswordService setChangePasswordService(string serviceUrl, string username, string password, string nonce)
        {
            ChangePasswordService oChangePasswordService = new ChangePasswordService();
            oChangePasswordService.SetUrl(serviceUrl);
            oChangePasswordService.SetUsernameToken(username, password, nonce);
            
            return oChangePasswordService;
        }
    }
}
