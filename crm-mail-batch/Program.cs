using System;
using System.IO;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Microsoft.Extensions.Configuration;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Net.Http.Headers;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.DependencyCollector;
using Microsoft.ApplicationInsights.Extensibility;
using System.Net.Http;


namespace Saorsa.Outlook.Mail
{
    class Program
    {
        static async Task Main(string[] args){
            TelemetryConfiguration configuration = TelemetryConfiguration.CreateDefault();
            configuration.InstrumentationKey = "ENTERTELEMETRYID";
            TelemetryClient telemetryClient = new TelemetryClient(configuration);
            try{
                var appInformation = LoadApplicationInformation();
                await Program.CreateAuthorizationProvider(appInformation, telemetryClient);
            }
            catch(Exception ex){
                telemetryClient.TrackException(ex);
            }
            finally{
                telemetryClient.Flush();
            }
        }

        private static IConfigurationRoot LoadApplicationInformation(){

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
        private static string[] LoadUserIds(string fileName){
            string path;
            string fileContent;

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
        private static async Task CreateAuthorizationProvider(IConfigurationRoot appInformation, TelemetryClient telemetryClient){
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
            var userIds = LoadUserIds(appInformation["userFileName"]);
            List<QueryOption> options = new List<QueryOption>{
                new QueryOption("$search", searchParam)
            };
            List<UserEmailBox> emails = new List<UserEmailBox>();
            foreach(var item in userIds){
                var mailMessages = await graphServiceClient.Users[item].Messages.Request(options).Select("subject, sender, body, isRead, conversationId, id").GetAsync();
                UserEmailBox userEmails = new UserEmailBox();
                userEmails.UserId = item;
                foreach(var mail in mailMessages){
                    if(mail.IsRead == false){
                        telemetryClient.TrackTrace($"I got unread item for User with ID: {item} for Item ID: {mail.Id} Conversation ID: {mail.ConversationId} with Subject: {mail.Subject}");
                        EmailMessage message = new EmailMessage();
                        message.Subject = mail.Subject;
                        message.SetSenderEmail(mail.Sender.EmailAddress);
                        message.Message = mail.Body.Content;
                        message.ItemId = mail.Id;
                        message.ConversationId = mail.ConversationId;
                        userEmails.SetUserEmail(message);
                    }
                }
                emails.Add(userEmails);
            }

        }
    }
}