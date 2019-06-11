using System.Net.Http;
using System.Net.Http.Headers;

namespace BDO_Localisation_AddOn.BOG_Integration_Services
{
    public static class HttpClientExtensions
    {
        public static void SetBearerToken(this HttpClient client, string token)
        {
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        }
    }
}