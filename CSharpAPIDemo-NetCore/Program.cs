using Microsoft.Extensions.Configuration;
using Microsoft.PowerPlatform.Dataverse.Client;
using System;
using System.IO;

namespace CSharpAPIDemo_NetCore
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Retrieve the required values from Secrets.json
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("Secrets.json", optional: false);

            IConfiguration config = builder.Build();

            var mySecretValues = config.GetSection("MySecretValues").Get<MySecretValues>();

            string url = mySecretValues.Url;
            string clientId = mySecretValues.ClientId;
            string clientSecret = mySecretValues.ClientSecret;
            string connectString = $"AuthType=ClientSecret;url={mySecretValues.Url};ClientId={mySecretValues.ClientId};ClientSecret={mySecretValues.ClientSecret}";
            string authority = mySecretValues.Authority;

            using (var svc = new ServiceClient(connectString))
            {
              
            }

            Console.WriteLine("Hello World!");
        }
    }

    // Used to access Secrets.json
    public class MySecretValues
    {
        public string Url { get; set; }
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string TenantId { get; set; }
        public string Authority { get; set; }
    }
}
