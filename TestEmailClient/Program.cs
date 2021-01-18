using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Converters;
using System;
using System.Globalization;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace TestEmailClient
{
    class Program
    {
        private const string addInstance = "https://login.microsoftonline.com/{0}";
        private const string resource = "https://graph.microsoft.com";

        //Personal Domain Details
        private const string clientId = "d438bebc-42ea-4230-9622-dba9349eed3e";
        private const string tenant = "testTenantKM.onmicrosoft.com";
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
            Task<string> usersList = GetUsers(token.Result);
            usersList.Wait();
            Console.WriteLine("Users Details : ");
            Console.WriteLine(usersList.Result);
            JObject usersListResponse = JObject.Parse(usersList.Result);
            Console.WriteLine("usersListResponse : " + usersListResponse.ToString());
            JArray usersListArray = (JArray)usersListResponse["value"];

            foreach (var userDetail in usersListArray)
            {
                var user = Newtonsoft.Json.JsonConvert.DeserializeObject<User>(userDetail.ToString());
                Console.WriteLine(user.ToString());
                Console.WriteLine("Fetching Details for User");

                Task<string> userDetails = GetUserDetails(user, token.Result);
                userDetails.Wait();
                Console.WriteLine("User Details : " + userDetails.Result);

                Task<string> userMailDetails = GetUserMails(token.Result);
                userMailDetails.Wait();
                Console.WriteLine("Emails : " + userMailDetails.Result);
            }
        }

        private static async Task<string> GetUserDetails(User user, string result)
        {
            string userDetails = null;
            var uri = "https://graph.microsoft.com/v1.0/" + tenetId + "/users/" + user.id;
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

    class User
    {
        public string[] businessPhones { get; set; }
        public string displayName { get; set; }
        public string givenName { get; set; }
        public string jobTitle { get; set; }
        public string mail { get; set; }
        public string mobilePhone { get; set; }
        public string officeLocation { get; set; }
        public string preferredLanguage { get; set; }
        public string surName { get; set; }
        public string userPrincipalName { get; set; }
        public string id { get; set; }

        public override string ToString()
        {
            string toStr = "User Details";
            toStr = toStr + "\n" + "ID :" + "\t" + id;
            toStr = toStr + "\n" + "Given Name :" + "\t" + givenName;
            toStr = toStr + "\n" + "Display Name :" + "\t" + displayName;
            toStr = toStr + "\n" + "Surname :" + "\t" + surName;
            toStr = toStr + "\n" + "User Principal Name :" + "\t" + userPrincipalName;
            toStr = toStr + "\n" + "Job Title :" + "\t" + jobTitle;
            toStr = toStr + "\n" + "Mail :" + "\t" + mail;
            toStr = toStr + "\n" + "Mobile Phone :" + "\t" + mobilePhone;
            toStr = toStr + "\n" + "Office Location :" + "\t" + officeLocation;            
            toStr = toStr + "\n" + "Business Phones :" + "\n";
            
            foreach (var v in businessPhones) {
                toStr = toStr  + "\t" + v + "\n";
            }
            return toStr;
        }

    }
}
