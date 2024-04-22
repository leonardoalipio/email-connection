using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace Email.Office365
{
    public class ApiHandler
    {
        public static void InitializeClient(HttpClient ApiClient)
        {
            if (ApiClient == null)
                ApiClient = new HttpClient();

            ApiClient.DefaultRequestHeaders.Accept.Clear();
            ApiClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"));
            ApiClient.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64)");
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11;
        }

        public static async Task<AccessTokenModel> GetAccessTokenAsync(string client_secret, string client_id, string tenant_id)
        {
            string url = $"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token";

            var data = new Dictionary<string, string>
            {
                {"grant_type", "client_credentials"},
                {"scope", "https://outlook.office365.com/.default"},
                {"client_id",  client_id},
                {"client_secret", client_secret}
            };

            using (var httpClient = new HttpClient())
            {
                InitializeClient(httpClient);
                var response = await httpClient.PostAsync(url, new FormUrlEncodedContent(data));
                return await response.Content.ReadAsAsync<AccessTokenModel>();
            }
        }
    }
}
