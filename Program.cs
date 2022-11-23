using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROM.TSIS2.CSharpAPIDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Retrieve the required values from Secrets.config
            string url = ConfigurationManager.AppSettings["Url"];
            string clientId = ConfigurationManager.AppSettings["ClientId"];
            string clientSecret = ConfigurationManager.AppSettings["ClientSecret"];
            string connectString = $"AuthType=ClientSecret;url={url};ClientId={clientId};ClientSecret={clientSecret}";

            // Connect to the TSIS 2 - ROM API
            using (var svc = new CrmServiceClient(connectString))
            {
                // EXAMPLE - Retrieve all Regions in English
                {
                    EntityCollection regions = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.All_Regions_English));

                    Console.WriteLine("All Regions in English:");
                    Console.WriteLine();

                    foreach (var region in regions.Entities)
                    {
                        Console.WriteLine($"ID: {region.Id.ToString()}");
                        Console.WriteLine($"Name: {region.GetAttributeValue<String>("ovs_territorynameenglish")}");
                        Console.WriteLine();
                    }
                }

                // EXAMPLE - Retrieve all Regions in French
                {
                    EntityCollection regions = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.All_Regions_French));

                    Console.WriteLine("All Regions in French:");
                    Console.WriteLine();

                    foreach (var region in regions.Entities)
                    {
                        Console.WriteLine($"ID: {region.Id.ToString()}");
                        Console.WriteLine($"Name: {region.GetAttributeValue<String>("ovs_territorynamefrench")}");
                        Console.WriteLine();
                    }
                }

                // EXAMPLE - Retrieve all Aviation Security Stakeholders
                {
                    EntityCollection stakeholders = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.All_Operational_AvSec_Stakeholders));

                    Console.WriteLine("All AvSec Stakeholders:");
                    Console.WriteLine();

                    foreach (var stakeholder in stakeholders.Entities)
                    {
                        Console.WriteLine($"ID: {stakeholder.Id.ToString()}");
                        Console.WriteLine($"Name: {stakeholder.GetAttributeValue<String>("name")}");
                        Console.WriteLine();
                    }
                }

                // EXAMPLE - Retrieve all Intermodal Surface Security Oversight Stakeholders
                {
                    EntityCollection stakeholders = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.All_Operational_ISSO_Stakeholders));

                    Console.WriteLine("All ISSO Stakeholders:");
                    Console.WriteLine();

                    foreach (var stakeholder in stakeholders.Entities)
                    {
                        Console.WriteLine($"ID: {stakeholder.Id.ToString()}");
                        Console.WriteLine($"Name: {stakeholder.GetAttributeValue<String>("name")}");
                        Console.WriteLine();
                    }
                }
            }

            Console.Read();
        }
    }
}
