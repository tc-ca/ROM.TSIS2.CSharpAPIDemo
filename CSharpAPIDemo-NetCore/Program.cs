using Microsoft.Extensions.Configuration;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
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

            // Connect to the TSIS 2 - ROM API using the ServiceClient
            using (var svc = new ServiceClient(connectString))
            {
                // EXAMPLE - Retrieve all Security Incident Types
                {
                    // English
                    EntityCollection securityIncidentTypesEnglish = svc.RetrieveMultiple(new FetchExpression(LookupFetchXML.SecurityIncidents()));

                    // French
                    EntityCollection securityIncidentTypesFrench = svc.RetrieveMultiple(new FetchExpression(LookupFetchXML.SecurityIncidents(false)));

                    foreach (var securityIncident in securityIncidentTypesEnglish.Entities)
                    {
                        Console.WriteLine($"ID: {securityIncident.Id.ToString()}");
                        Console.WriteLine($"Name: {securityIncident.GetAttributeValue<String>(LookupColumns.SecurityIncidentType.EnglishName)}");
                    }
                }
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
        public string Authority { get; set; }
    }
}
