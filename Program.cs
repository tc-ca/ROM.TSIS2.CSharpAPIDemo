using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace ROM.TSIS2.CSharpAPIDemo
{
    internal class Program
    {
        // Retrieve the required values from Secrets.config
        static string url = ConfigurationManager.AppSettings["Url"];
        static string clientId = ConfigurationManager.AppSettings["ClientId"];
        static string clientSecret = ConfigurationManager.AppSettings["ClientSecret"];
        static string connectString = $"AuthType=ClientSecret;url={url};ClientId={clientId};ClientSecret={clientSecret}";
        static string authority = ConfigurationManager.AppSettings["Authority"];

        // API path for REST calls
        static string apiPath = "/api/data/v9.2/";

        static void Main(string[] args)
        {
            // Connect to the TSIS 2 - ROM API using the HttpClient
            Task.WaitAll(Task.Run(async () => await GetTableData()));

            // Connect to the TSIS 2 - ROM API using the CrmServiceClient
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

                // EXAMPLE - Retrieve all location types
                var locationTypes = GetChoicesExample.GetChoices(svc.GetGlobalOptionSetMetadata("ts_locationtype"));

                // EXAMPLE - Create a Security Incident
                {
                    Entity newSecurityIncident = new Entity("ts_securityincident");

                    // Get a mode
                    var modes = GetChoicesExample.GetChoices(svc.GetGlobalOptionSetMetadata("ts_securityincidentmode"));
                    var selectedMode = modes[0];

                    // Set the mode for the new Security Incident
                    newSecurityIncident.Attributes["ts_mode"] = new OptionSetValue(Convert.ToInt32(selectedMode.Value));

                    // Get a Reporting Company
                    EntityCollection reportingCompanies = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.All_Operational_AvSec_Stakeholders));
                    var selectedReportingCompany = reportingCompanies[0].Id;

                    // Set the reporting company for the new Security Incident
                    newSecurityIncident.Attributes["ts_stakeholder"] = new EntityReference("account", Guid.Parse(selectedReportingCompany.ToString()));

                    // Record the ID (GUID) of the new Security Incident
                    var newSecurityIncidentID = svc.Create(newSecurityIncident);
                }

                // EXAMPLE - Update a Security Incident
                {
                    // Get the record you want to update
                    EntityCollection securityIncidents = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.All_Active_SecurityIncidents));
                    var selectedSecurityIncident = securityIncidents[0].Id;

                    var updateSecurityIncident = new Entity("ts_securityincident", selectedSecurityIncident);

                    // Update a field with some random text
                    updateSecurityIncident.Attributes["ts_arrestsdetails"] = "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean";

                    svc.Update(updateSecurityIncident);
                }

                // EXAMPLE - Retrieve all column names of a table
                {
                    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                    {
                        EntityFilters = EntityFilters.All,
                        LogicalName = "msdyn_functionallocation"
                    };

                    RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)svc.Execute(retrieveEntityRequest);
                    EntityMetadata entity = retrieveAccountEntityResponse.EntityMetadata;

                    Console.WriteLine( $"{entity.SchemaName} - {entity.DisplayName.UserLocalizedLabel.Label} entity meta-data:");

                    Console.WriteLine("Entity attributes:");
                    foreach (object attribute in entity.Attributes)
                    {
                        AttributeMetadata a = (AttributeMetadata)attribute;
                        Console.WriteLine(a.LogicalName);
                    }
                }
            }

            Console.Read();
        }

        private static async Task GetTableData()
        {
            // Authenticate with the API
            AuthenticationResult _authResult;
            AuthenticationContext authenticationContext = new AuthenticationContext(authority);
            ClientCredential credential = new ClientCredential(clientId, clientSecret);
            _authResult = await authenticationContext.AcquireTokenAsync(url, credential);

            using (HttpClient httpClient = new HttpClient())
            {
                httpClient.BaseAddress = new Uri(url);
                httpClient.Timeout = new TimeSpan(0, 2, 0);
                httpClient.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
                httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");
                httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _authResult.AccessToken);

                // Make the GET call
                // NOTE: Documentation on how to query data can be found here:
                // https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query-data-web-api

                // Example of a filter being applied while we get all the sites
                var response = httpClient.GetAsync(apiPath + "msdyn_functionallocations?$select=msdyn_functionallocationid,msdyn_name&$filter=ts_sitestatus eq 717750000&$top=5").Result;

                if (response.IsSuccessStatusCode)
                {
                    var jRetrieveResponse = JObject.Parse(response.Content.ReadAsStringAsync().Result);
                    dynamic collAccounts = JsonConvert.DeserializeObject(jRetrieveResponse.ToString());

                    foreach (var data in collAccounts.value)
                    {
                        // Example of writing out the output
                        Console.WriteLine("Site ID: " + data.msdyn_functionallocationid);
                        Console.WriteLine("Site Name: " + data.msdyn_name);
                        Console.WriteLine("");
                    }
                }
                else
                {
                    return;
                }
            }
        }

    }
}
