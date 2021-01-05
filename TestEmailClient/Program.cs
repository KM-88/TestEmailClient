using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json.Linq;
using System;
using System.Globalization;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace TestEmailClient
{
    class Program
    {
        private const string clientId = "d438bebc-42ea-4230-9622-dba9349eed3e";
        private const string addInstance = "https://login.microsoftonline.com/{0}";
        private const string tenant = "testTenantKM.onmicrosoft.com";
        private const string resource = "https://graph.microsoft.com";
        private const string appKey = "r~Ap~zklDe_~2ceHW.T3_z16wcg3Ovt_88";
        private const string tenetId = "36e750b2-01a3-4d84-903a-7349310f1cb1";
        static string authority = String.Format(CultureInfo.InvariantCulture, addInstance, tenant);

        private static HttpClient httpClient = new HttpClient();
        private static AuthenticationContext context = null;
        private static ClientCredential credential = null;

        static void Main(string[] args)
        {
            context = new AuthenticationContext(authority);
            credential = new ClientCredential(clientId, appKey);

            Task<string> token = GetToken();
            token.Wait();
            Console.WriteLine(token.Result);
            Task<string> users = GetUsers(token.Result);
            users.Wait();
            Console.WriteLine(users.Result);
            JObject joResponse = JObject.Parse(users.Result);
            JArray array = (JArray)joResponse["value"];

            Console.WriteLine("ID : " + array[0]);

            Task<string> userDetails = GetUserDetails(token.Result);
            userDetails.Wait();
            Console.WriteLine("User Details : " + userDetails.Result);

            Task<string> userMailDetails = GetUserMails(token.Result);
            userMailDetails.Wait();
            Console.WriteLine("Emails : " + userMailDetails.Result);
        }

        private static async Task<string> GetUserDetails(string result) {
            string id = "850bc806-38af-48ce-a6dd-62338cbbe0a2";
            string userDetails = null;
            var uri = "https://graph.microsoft.com/v1.0/" + tenetId + "/users/"+ id;
            Console.WriteLine("URI : " + uri);
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result);
            var getUsersResult = await httpClient.GetAsync(uri);
            if (getUsersResult.Content != null)
            {
                userDetails = await getUsersResult.Content.ReadAsStringAsync();
            }
            return userDetails;
        }

        private static async Task<string> GetUserMails(string result)
        {
            string id = "850bc806-38af-48ce-a6dd-62338cbbe0a2";
            string userDetails = null;
            var uri = "https://graph.microsoft.com/v1.0/" + tenetId + "/users/" + id + "/messages";
            Console.WriteLine("URI : " + uri);
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result);
            var getUsersResult = await httpClient.GetAsync(uri);
            if (getUsersResult.Content != null)
            {
                userDetails = await getUsersResult.Content.ReadAsStringAsync();
            }
            return userDetails;
        }


        private static async Task<string> GetUsers(string result)
        {
            string users = null;
            var uri = "https://graph.microsoft.com/v1.0/" + tenetId + "/users";
            Console.WriteLine("URI : " + uri);
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result);
            var getUsersResult = await httpClient.GetAsync(uri);
            if (getUsersResult.Content != null)
            {
                users = await getUsersResult.Content.ReadAsStringAsync();
            }
            return users;
        }

        private static async Task<string> GetToken()
        {
            AuthenticationResult authenticationResult = null;
            string token = null;
            authenticationResult = await context.AcquireTokenAsync(resource, credential);
            token = authenticationResult.AccessToken;
            return token;
        }
    }
}
