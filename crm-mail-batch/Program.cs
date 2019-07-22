using System;
using System.IO;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Microsoft.Extensions.Configuration;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Net.Http.Headers;
using Newtonsoft.Json;

namespace crm_mail_batch
{
    class Program
    {
        static async Task Main(string[] args){
            
            var appInformation = LoadApplicationInformation();
            await Program.CreateAuthorizationProvider(appInformation);

        }

        private static IConfigurationRoot LoadApplicationInformation(){
            try{
                var appInformation = new ConfigurationBuilder()
                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                .AddJsonFile("appsetting.json", false, true)
                .Build();

                if(string.IsNullOrEmpty(appInformation["applicationId"]) || string.IsNullOrEmpty(appInformation["applicationSecret"])|| string.IsNullOrEmpty(appInformation["tenantId"])|| string.IsNullOrEmpty(appInformation["redirectUri"]) || string.IsNullOrEmpty("userFileName") || string.IsNullOrEmpty("searchParam")){
                        Console.WriteLine("Missing information in appsetting.json");
                        return null;
                    }
                return appInformation;
                }
            catch (System.IO.FileNotFoundException){
                
                return null;

                }
            }
        private static string[] LoadUserId(string fileName){
            try{
                string path;
                string fileContent;

                //extend to handle windows dirs
                //var os = Environment.OSVersion.Platform;

                if(fileName.Contains(".txt")){
                    path = Path.Combine(System.IO.Directory.GetCurrentDirectory(), fileName);
                }
                else{
                    path = Path.Combine(System.IO.Directory.GetCurrentDirectory(), fileName + ".txt");
                }
                fileContent = System.IO.File.ReadAllText(path);
                if(string.IsNullOrEmpty(fileContent)){
                    Console.WriteLine("No user ID/s or userPrincipalName/s in the file");
                    return null;
                }
                return fileContent.Split(",");
            }
            catch(System.IO.FileNotFoundException){
                return null;
                }
            }
        private static async Task CreateAuthorizationProvider(IConfigurationRoot appInformation){
            var clientId = appInformation["applicationId"];
            var tenantId = appInformation["tenantId"];
            var clientSecret = appInformation["applicationSecret"];
            var redirectUri = appInformation["redirectUri"];
            var searchParam = appInformation["searchParam"];
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithAuthority(authorityUri: $"https://login.microsoftonline.com/{tenantId}/v2.0")
                .WithClientSecret(clientSecret)
                .WithRedirectUri(redirectUri)
                .Build();

            ClientCredentialProvider authenticationProvider = new ClientCredentialProvider(confidentialClientApplication);
            string[] scopes = new string[] { "https://graph.microsoft.com/.default" };

            GraphServiceClient graphServiceClient =
    new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) => {
            var authResult = await confidentialClientApplication
                .AcquireTokenForClient(scopes)
                .ExecuteAsync();
            requestMessage.Headers.Authorization = 
                new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                })
            );
            var userIds = LoadUserId(appInformation["userFileName"]);
            List<QueryOption> options = new List<QueryOption>{
                new QueryOption("$search", searchParam)
            };
            List<UserEmailBox> emails = new List<UserEmailBox>();
            foreach(var item in userIds){
                var mailMessages = await graphServiceClient.Users[item].Messages.Request(options).Select("subject, sender, body, isRead").GetAsync();
                UserEmailBox userEmails = new UserEmailBox();
                userEmails.UserId = item;
                foreach(var mail in mailMessages){
                    if(mail.IsRead == false){
                        EmailMessage message = new EmailMessage();
                        message.Subject = mail.Subject;
                        message.SetSenderEmail(mail.Sender.EmailAddress);
                        message.Message = mail.Body.Content;
                        userEmails.SetUserEmail(message);
                    }
                }
                emails.Add(userEmails);
            }

            string json = JsonConvert.SerializeObject(emails);
            Console.Write(json);

        }

    }
}