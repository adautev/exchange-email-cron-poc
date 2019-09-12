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
using Microsoft.ApplicationInsights.Extensibility;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using Newtonsoft.Json;
using System.Text;
using Serilog;

namespace Saorsa.Outlook.Mail
{
    class Program
    {
        private static GraphServiceClient _graphServiceClient;
        private static TelemetryClient _telemetryClient;
        private static readonly HttpClient _client = new HttpClient();
        private static readonly List<EmailMessage> _emails =  new List<EmailMessage>();

        static async Task Main(string[] args)
        {
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console()
                .WriteTo.File("logs\\parser.log", rollingInterval: RollingInterval.Day)
                .CreateLogger();
            var appInformation = LoadApplicationInformation();
            TelemetryConfiguration configuration = TelemetryConfiguration.CreateDefault();
            configuration.InstrumentationKey = appInformation["telemetryId"];
            _telemetryClient = new TelemetryClient(configuration);            
            CreateAuthorizationProvider(appInformation);
            try {
               await Parse(appInformation);
               await SendData(appInformation);
            } catch (Exception ex)  {
                Log.Error(ex,"General exception occured.");
                _telemetryClient.TrackException(ex);
            } finally {
                _telemetryClient.Flush();
            }
        }

        private static IConfigurationRoot LoadApplicationInformation()
        {

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
        private static string[] LoadUserIds(string fileName)
        {
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
                Log.Error("No user ID/s or userPrincipalName/s in the file");
                return null;
            }
            return fileContent.Split(",");
        }
        
        private static async Task Parse(IConfigurationRoot appInformation) 
        {
            var searchParam = appInformation["searchParam"];
            var userIds = LoadUserIds(appInformation["userFileName"]);
            List<QueryOption> options = new List<QueryOption>{
                new QueryOption("$search", searchParam)
            };
            foreach(var item in userIds){
                if (String.IsNullOrEmpty(item)) {
                    continue;
                }
                var mailMessages = await _graphServiceClient.Users[item].Messages.Request(options).Select("subject, sender, body, isRead, conversationId, id").GetAsync();
                foreach(var mail in mailMessages){
                    if(!mail.IsRead.Value){
                        string messageTemplate = $"I got unread item for User with ID: {item} for Item ID: {mail.Id} Conversation ID: {mail.ConversationId} with Subject: {mail.Subject}";
                        _telemetryClient.TrackTrace(messageTemplate);
                        EmailMessage message = new EmailMessage(appInformation["searchParam"]);
                        message.Subject = mail.Subject;
                        message.SenderEmail = mail.Sender.EmailAddress;
                        message.Message = mail.Body.Content;
                        message.ItemId = mail.Id;
                        message.ConversationId = mail.ConversationId;
                        _emails.Add(message);
                    }
                }
                
            }
        }

        private static async Task SendData(IConfigurationRoot appInformation)
        {
            byte[] data = System.Text.ASCIIEncoding.ASCII.GetBytes(String.Concat(appInformation["SNOWUsername"],":",appInformation["SNOWPassword"]));
            String authHeader = System.Convert.ToBase64String(data);
            foreach (var email in _emails)
            {          
                var serializer = new DataContractJsonSerializer(typeof(EmailMessage));
                _client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("basic", String.Concat(authHeader));
                string serializedContent = JsonConvert.SerializeObject(email);
                Log.Debug($"Sending {serializedContent}");
                var serializedEmail = new StringContent(serializedContent, Encoding.UTF8, "application/json");
                var response = await _client.PostAsync(appInformation["SNOWURL"], serializedEmail);
                Log.Debug($"Got response {response}");
            }
        }

        private static void CreateAuthorizationProvider(IConfigurationRoot appInformation) 
        {
            var clientId = appInformation["applicationId"];
            var tenantId = appInformation["tenantId"];
            var clientSecret = appInformation["applicationSecret"];
            var redirectUri = appInformation["redirectUri"];
            
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithAuthority(authorityUri: $"https://login.microsoftonline.com/{tenantId}/v2.0")
                .WithClientSecret(clientSecret)
                .WithRedirectUri(redirectUri)
                .Build();

            ClientCredentialProvider authenticationProvider = new ClientCredentialProvider(confidentialClientApplication);
            string[] scopes = new string[] { "https://graph.microsoft.com/.default" };
            
            _graphServiceClient =
            new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) => {
            var authResult = await confidentialClientApplication
                .AcquireTokenForClient(scopes)
                .ExecuteAsync();
                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                })
            );            
        }
    }
}