using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Net.Http;

namespace GraphDemo
{
    class Program
    {
        //application (client) ID
        //451912ea-1728-4fc5-b27d-453821c61aa6
        //Directory (tenant) ID
        //c10ade6e-bb7a-4ffc-bad2-b86a07108640
        //client secret 
        //u(6OCN;krp_8)+ZVs2QqOsp=qtclv-Txd10vl[;
        //redirect URI
        //https://localhost:8080

        private static GraphServiceClient _graphServiceClient;
        private static HttpClient _httpClient;

        private const string alias = "foxdave";

        static void Main(string[] args)
        {
            //Console.WriteLine("Hello World!");
            var config = LoadAppSettings();
            if (null == config)
            {
                Console.WriteLine("Missing or invalid appsettings.json file. Please see README.md for configuration instructions.");
                return;
            }

            //Query using Graph SDK (preferred when possible)
            GraphServiceClient graphClient = GetAuthenticatedGraphClient(config);

            //Direct query using HTTPClient (for beta endpoint calls or not available in Graph SDK)
            HttpClient httpClient = GetAuthenticatedHTTPClient(config);

            #region Day 18
            //获取当前时区设置
            GetUserMailboxDefaultTimeZone();

            //更新当前用户邮箱的时区设置
            SetUserMailboxDefaultTimeZone();

            //再次获取时区设置验证更新是否成功
            GetUserMailboxDefaultTimeZone();

            //通过MS Graph SDK获取邮件消息
            ListUserMailInboxMessages();

            //创建一个新的消息规则
            CreateUserMailBoxRule();

            //获取消息规则以验证创建是否成功
            ListUserMailBoxRules();

            Console.ReadKey();
            #endregion

            //为用户分配license
            AddLicenseToUser(config);

            //在Azure AD中创建用户
            CreateAndFindNewUser(config);

            Console.WriteLine("User Created.");
            Console.ReadKey();


            //在.NET Core应用程序中调用Microsoft Graph获取Office 365用户信息
            List<QueryOption> options = new List<QueryOption>
            {
                new QueryOption("$top", "1")
            };

            var graphResult = graphClient.Users.Request(options).GetAsync().Result;
            Console.WriteLine("Graph SDK Result");
            Console.WriteLine(graphResult[0].DisplayName);

            Uri Uri = new Uri("https://graph.microsoft.com/v1.0/users?$top=1");
            var httpResult = httpClient.GetStringAsync(Uri).Result;

            Console.WriteLine("HTTP Result");
            Console.WriteLine(httpResult);

            Console.ReadKey();
        }

        private static void ListUserMailBoxRules()
        {
            var mailboxHelper = new MailboxHelper(_graphServiceClient);
            List<ResultsItem> rules = mailboxHelper.GetUserMailboxRules(alias).Result;
            Console.WriteLine("Rules count: " + rules.Count);
            foreach (ResultsItem rule in rules)
            {
                Console.WriteLine("Rule Name: " + rule.Display);
            }
        }

        private static void CreateUserMailBoxRule()
        {
            var mailboxHelper = new MailboxHelper(_graphServiceClient);
            mailboxHelper.CreateRule(alias, "ForwardBasedonSender", 2, true, "svarukal", "adelev@M365x995052.onmicrosoft.com").GetAwaiter().GetResult();
        }

        private static void GetUserMailboxDefaultTimeZone()
        {
            var mailboxHelper = new MailboxHelper(_graphServiceClient);
            var defaultTimeZone = mailboxHelper.GetUserMailboxDefaultTimeZone(alias).Result;
            Console.WriteLine("Default timezone: " + defaultTimeZone);
        }
        private static void SetUserMailboxDefaultTimeZone()
        {
            var mailboxHelper = new MailboxHelper(_graphServiceClient, _httpClient);
            mailboxHelper.SetUserMailboxDefaultTimeZone(alias, "China Standard Time");
        }
        private static void ListUserMailInboxMessages()
        {
            var mailboxHelper = new MailboxHelper(_graphServiceClient);
            List<ResultsItem> items = mailboxHelper.ListInboxMessages(alias).Result;
            Console.WriteLine("Message count: " + items.Count);
            foreach (ResultsItem item in items)
            {
                Console.WriteLine(item.Display);
            }
        }

        private static void AddLicenseToUser(IConfigurationRoot config)
        {
            string domain = config["domain"];
            string upn = $"{alias}@{domain}";

            var userHelper = new UserHelper(_graphServiceClient);
            var user = userHelper.GetUser(upn).Result;

            var licenseHelper = new LicenseHelper(_graphServiceClient);
            var sku = licenseHelper.GetLicense().Result;
            licenseHelper.AddLicense(user.Id, sku.SkuId).GetAwaiter().GetResult();
            Console.WriteLine("License assigned.");
        }

        private static IConfigurationRoot LoadAppSettings()
        {
            try
            {
                var config = new ConfigurationBuilder()
                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", false, true)
                .Build();

                // Validate required settings
                if (string.IsNullOrEmpty(config["applicationId"]) ||
                string.IsNullOrEmpty(config["applicationSecret"]) ||
                string.IsNullOrEmpty(config["redirectUri"]) ||
                string.IsNullOrEmpty(config["tenantId"]) ||
                string.IsNullOrEmpty(config["domain"]))
                {
                    return null;
                }

                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var clientId = config["applicationId"];
            var clientSecret = config["applicationSecret"];
            var redirectUri = config["redirectUri"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = new ConfidentialClientApplication(clientId, authority, redirectUri, new ClientCredential(clientSecret), null, null);
            return new MsalAuthenticationProvider(cca, scopes.ToArray());
        }

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _graphServiceClient = new GraphServiceClient(authenticationProvider);
            return _graphServiceClient;
        }

        private static HttpClient GetAuthenticatedHTTPClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _httpClient = new HttpClient(new AuthHandler(authenticationProvider, new HttpClientHandler()));
            return _httpClient;
        }

        private static void CreateAndFindNewUser(IConfigurationRoot config)
        {
            const string alias = "sdk_test";
            string domain = config["domain"];
            var userHelper = new UserHelper(_graphServiceClient);
            userHelper.CreateUser("SDK Test User", alias, domain, "ChangeThis!0").GetAwaiter().GetResult();
            var user = userHelper.FindByAlias(alias).Result;
            // Console writes for demo purposes
            Console.WriteLine(user.DisplayName);
            Console.WriteLine(user.UserPrincipalName);
        }
    }
}
