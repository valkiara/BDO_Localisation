using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace BDO_Localisation_AddOn.TBC_Integration_Services
{
    static partial class MainMovementService
    {
        public static BaseQueryResultIo getAccountMovements(MovementService oMovementService, AccountMovementFilterIo oAccountMovementFilterIo, out AccountMovementDetailIo[] oAccountMovementDetailIo, out string errorText)
        {
            errorText = null;
            oAccountMovementDetailIo = null;
            BaseQueryResultIo oBaseQueryResultIo = null;

            try
            {
                oBaseQueryResultIo = oMovementService.GetAccountMovements(oAccountMovementFilterIo, out oAccountMovementDetailIo);
            }
            catch (Exception ex)
            {
                try
                {
                    errorText = ex.Message + '\n' + ((System.Web.Services.Protocols.SoapException)ex).Code.Name;
                    return oBaseQueryResultIo;
                }
                catch
                {
                    errorText = ex.Message;
                    try
                    {
                        if (errorText.Contains("<title>"))
                        {
                            errorText = BDOSResources.getTranslate("WebsiteIsUnderConstruction"); //"მიმდინარეობს ტექნიკური სამუშაოები";
                        }
                    }
                    catch { }

                    return oBaseQueryResultIo;
                }
            }           
            return oBaseQueryResultIo;
        }

        /// <summary>ავტორიზაციის პარამეტრების შევსება</summary>
        /// <param name="serviceUrl"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="nonce"></param>
        /// <returns></returns>
        public static MovementService setMovementService(string serviceUrl, string username, string password, string nonce)
        {
            MovementService oMovementService = new MovementService();
            oMovementService.SetUrl(serviceUrl);
            oMovementService.SetUsernameToken(username, password, nonce);

            return oMovementService;
        }
    }
}
