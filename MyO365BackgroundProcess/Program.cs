using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace MyO365BackgroundProcess
{
    class Program
    {
        private static string CLIENT_ID = "4b7fb8dd-0b22-45a2-8248-3cc87a3560a7";
        private static string PRIVATE_KEY_PASSWORD = "P@ssword"; //THIS IS BAD...USE AZURE KEY VAULT
        static void Main(string[] args)
        {
            doStuffInOffice365().Wait();
        }

        private async static Task doStuffInOffice365()
        {
            //set the authentication context
            //you can do multi-tenant app-only, but you cannot use /common for authority...must get tenant ID
            string authority = "https://login.windows.net/rzna.onmicrosoft.com/";
            AuthenticationContext authenticationContext = new AuthenticationContext(authority, false);

            //read the certificate private key from the executing location
            //NOTE: This is a hack...Azure Key Vault is best approach
            var certPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            certPath = certPath.Substring(0, certPath.LastIndexOf('\\')) + "\\O365AppOnly_private.pfx";
            var certfile = System.IO.File.OpenRead(certPath);
            var certificateBytes = new byte[certfile.Length];
            certfile.Read(certificateBytes, 0, (int)certfile.Length);
            var cert = new X509Certificate2(
                certificateBytes,
                PRIVATE_KEY_PASSWORD,
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.PersistKeySet); //switchest are important to work in webjob
            ClientAssertionCertificate cac = new ClientAssertionCertificate(CLIENT_ID, cert);

            //get the access token to SharePoint using the ClientAssertionCertificate
            Console.WriteLine("Getting app-only access token to SharePoint Online");
            var authenticationResult = await authenticationContext.AcquireTokenAsync("https://rzna.sharepoint.com/", cac);
            var token = authenticationResult.AccessToken;
            Console.WriteLine("App-only access token retreived");

            //perform a post using the app-only access token to add SharePoint list item in Attendee list
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");

            //create the item payload for saving into SharePoint
            var itemPayload = new
            {
                __metadata = new { type = "SP.Data.SampleListItem" },
                Title = String.Format("Created at {0} {1} from app-only AAD token", DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString())
            };

            //setup the client post
            HttpContent content = new StringContent(JsonConvert.SerializeObject(itemPayload));
            content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
            Console.WriteLine("Posting ListItem to SharePoint Online");
            using (HttpResponseMessage response = await client.PostAsync("https://rzna.sharepoint.com/_api/web/Lists/getbytitle('Sample')/items", content))
            {
                if (!response.IsSuccessStatusCode)
                    Console.WriteLine("ERROR: SharePoint ListItem Creation Failed!");
                else
                    Console.WriteLine("SharePoint ListItem Created!");
            }
        }
    }
}
