using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Security;
using System.Linq;
using System.Collections.Generic;
using System.Text;

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

        // List to store records
        static List<FileItem>fileItems = new List<FileItem>();
        static List<FileItem>manyToManyFileItems = new List<FileItem>();
        static List<SharePointFileItem>sharePointFileItems = new List<SharePointFileItem>();
        //static List<Case>caseItems = new List<Case>();
        //static List<WorkOrder> workOrderItems = new List<WorkOrder>();
        //static List<WorkOrderServiceTask> workOrderServiceTaskItems = new List<WorkOrderServiceTask>();

        static async Task Main(string[] args)
        {
            // Connect to the TSIS 2 - ROM API using the HttpClient
            //Task.WaitAll(Task.Run(async () => await GetTableData()));

            int counter = 0;

            // Connect to the TSIS 2 - ROM API using the CrmServiceClient
            using (var svc = new CrmServiceClient(connectString))
            {
                // Get all the files
                {
                    EntityCollection files = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.All_Files));

                    // get all Cases,Work Orders, and Work Order Service Tasks
                    EntityCollection cases = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.All_Cases));
                    EntityCollection workOrders = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.All_WorkOrders));
                    EntityCollection workOrderServiceTasks = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.All_Files_WorkOrderServiceTask));

                    // put all Cases,Work Orders, and Work Order Service Tasks in a list
                    //{
                    //    foreach (var currentCase in cases.Entities)
                    //    {
                    //        caseItems.Add(new Case
                    //        {
                    //            CaseId = currentCase.GetAttributeValue<Guid>("incidentid").ToString()
                    //        });
                    //    }

                    //    foreach (var currentWorkOrder in workOrders.Entities)
                    //    {
                    //        workOrderItems.Add(new WorkOrder
                    //        {
                    //            WorkOrderId = currentWorkOrder.GetAttributeValue<Guid>("msdyn_workorderid").ToString(),
                    //            CaseId = currentWorkOrder.GetAttributeValue<Guid>("ts_incident").ToString()
                    //        });
                    //    }

                    //    foreach (var currentWorkOrderServiceTask in workOrderServiceTasks.Entities)
                    //    {
                    //        workOrderServiceTaskItems.Add(new WorkOrderServiceTask
                    //        {
                    //            WorkOrderId = currentWorkOrderServiceTask.GetAttributeValue<Guid>("msdyn_workorderid").ToString(),
                    //            WorkOrderServiceTaskId = currentWorkOrderServiceTask.GetAttributeValue<Guid>("msdyn_workorderservicetaskid").ToString()
                    //        });
                    //    }
                    //}

                    // get all individual file records
                    {
                        // put all the files in a list
                        foreach (var file in files.Entities)
                        {
                            string fileId = file.GetAttributeValue<Guid>("ts_fileid").ToString();
                            string fileName = DateTime.UtcNow.ToString("yyyyMMddHHmmssffff") + " " + file.GetAttributeValue<String>("ts_file");
                            bool? uploadedToSharePoint = file.GetAttributeValue<bool>("ts_uploadedtosharepoint");
                            Guid exemptionId = file.GetAttributeValue<EntityReference>("ts_exemption")?.Id ?? Guid.Empty;
                            Guid securityIncidentId = file.GetAttributeValue<EntityReference>("ts_securityincident")?.Id ?? Guid.Empty;

                            fileItems.Add(new FileItem
                            {
                                FileId = fileId,
                                FileName = fileName,
                                UploadedToSharePoint = uploadedToSharePoint,
                                ExemptionId = exemptionId,
                                SecurityIncidentId = securityIncidentId,
                            });
                        }
                    }

                    // now get all the many-to-many records for file
                    {
                        manyToManyFileItems.AddRange(GetFileItems(svc, FetchXMLExamples.All_Files_Cases, "incidentid", "ts_fileid"));
                        manyToManyFileItems.AddRange(GetFileItems(svc, FetchXMLExamples.All_Files_WorkOrders, "msdyn_workorderid", "ts_fileid"));
                        manyToManyFileItems.AddRange(GetFileItems(svc, FetchXMLExamples.All_Files_WorkOrderServiceTask, "msdyn_workorderservicetaskid", "ts_fileid"));
                        manyToManyFileItems.AddRange(GetFileItems(svc, FetchXMLExamples.All_Files_Findings, "ovs_findingid", "ts_fileid"));
                        manyToManyFileItems.AddRange(GetFileItems(svc, FetchXMLExamples.All_Files_Sites, "msdyn_functionallocationid", "ts_fileid"));
                        manyToManyFileItems.AddRange(GetFileItems(svc, FetchXMLExamples.All_Files_Operations, "ovs_operationid", "ts_fileid"));
                        manyToManyFileItems.AddRange(GetFileItems(svc, FetchXMLExamples.All_Files_Stakeholders, "accountid", "ts_fileid"));
                    }

                    // update the fileItems with the proper ID's
                    {
                        foreach (var fileItem in fileItems)
                        {
                            var itemIds = manyToManyFileItems.Where(x => x.FileId == fileItem.FileId);

                            foreach (var itemId in itemIds)
                            {
                                if (itemId.CaseId != Guid.Empty)
                                {
                                    fileItem.FileItemGroups.Add(new FileItemGroup{
                                        Id = itemId.CaseId,
                                        IdFieldName = itemId.FormIntegrationId
                                    });
                                }

                                if (itemId.WorkOrderId != Guid.Empty)
                                {
                                    fileItem.FileItemGroups.Add(new FileItemGroup
                                    {
                                        Id = itemId.WorkOrderId,
                                        IdFieldName = itemId.FormIntegrationId
                                    });
                                }

                                if (itemId.WorkOrderServiceTaskId != Guid.Empty)
                                {
                                    fileItem.FileItemGroups.Add(new FileItemGroup
                                    {
                                        Id = itemId.WorkOrderServiceTaskId,
                                        IdFieldName = itemId.FormIntegrationId
                                    });
                                }

                                if (itemId.StakeholderId != Guid.Empty)
                                {
                                    fileItem.FileItemGroups.Add(new FileItemGroup
                                    {
                                        Id = itemId.StakeholderId,
                                        IdFieldName = itemId.FormIntegrationId
                                    });
                                }

                                if (itemId.SiteId != Guid.Empty)
                                {
                                    fileItem.FileItemGroups.Add(new FileItemGroup
                                    {
                                        Id = itemId.SiteId,
                                        IdFieldName = itemId.FormIntegrationId
                                    });
                                }

                                if (itemId.OperationId != Guid.Empty)
                                {
                                    fileItem.FileItemGroups.Add(new FileItemGroup
                                    {
                                        Id = itemId.OperationId,
                                        IdFieldName = itemId.FormIntegrationId
                                    });
                                }
                            }
                        }
                    }

                    // do the SharePoint File logic 
                    {
                        foreach (var fileItem in fileItems.Where(x => x.UploadedToSharePoint == false && x.FileItemGroups.Count > 0))
                        {
                            foreach (var fileItemGroup in fileItem.FileItemGroups)
                            {
                                // Get the table name
                                string tableName = "";

                                if (fileItemGroup.IdFieldName == "incidentid")
                                    tableName = "Case";

                                if (fileItemGroup.IdFieldName == "msdyn_workorderid")
                                    tableName = "Work Order";

                                if (fileItemGroup.IdFieldName == "msdyn_workorderservicetaskid")
                                    tableName = "Work Order Service Task";

                                if (fileItemGroup.IdFieldName == "accountid")
                                    tableName = "Stakeholder";

                                if (fileItemGroup.IdFieldName == "msdyn_functionallocationid")
                                    tableName = "Site";

                                if (fileItemGroup.IdFieldName == "ovs_operationid")
                                    tableName = "Operation";

                                // Some Output
                                Console.WriteLine($"ID: {fileItem.FileId}");
                                Console.WriteLine($"File Name: {fileItem.FileName}");
                                Console.WriteLine($"{tableName}: {fileItemGroup.Id}");

                                // Define the columns you want to retrieve (null retrieves all columns)
                                ColumnSet columns = new ColumnSet(true);

                                Entity sharePointFile = null;

                                EntityCollection sharePointFiles = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.Single_SharePointFile(fileItemGroup.Id.ToString())));

                                sharePointFile = sharePointFiles.Entities.Count > 0 ? sharePointFiles.Entities[0] : null;

                                if (sharePointFile != null)
                                {
                                    fileItem.TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid");

                                    sharePointFileItems.Add(new SharePointFileItem
                                    {
                                        SharePointFileId = sharePointFile.Id.ToString(),
                                        SharePointFileGroupId = sharePointFile.GetAttributeValue<EntityReference>("ts_sharepointfilegroup")?.Id.ToString(),
                                        TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid"),
                                        TableName = sharePointFile.GetAttributeValue<string>("ts_tablename")
                                    });
                                }
                                else
                                {
                                    // Create the SharePoint File
                                    Entity newSharePointFile = new Entity("ts_sharepointfile");
                                    newSharePointFile.Attributes["ts_tablename"] = tableName;
                                    newSharePointFile.Attributes["ts_tablerecordid"] = fileItemGroup.Id.ToString();

                                    // Record the ID (GUID) of the new SharePoint File
                                    Guid createdSharePointFile = svc.Create(newSharePointFile);

                                    sharePointFiles = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.Single_SharePointFile(fileItemGroup.Id.ToString())));

                                    sharePointFile = sharePointFiles.Entities.Count > 0 ? sharePointFiles.Entities[0] : null;

                                    // Add the SharePoint File to the list
                                    sharePointFileItems.Add(new SharePointFileItem
                                    {
                                        SharePointFileId = sharePointFile.Id.ToString(),
                                        SharePointFileGroupId = sharePointFile.GetAttributeValue<EntityReference>("ts_sharepointfilegroup")?.Id.ToString(),
                                        TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid"),
                                        TableName = sharePointFile.GetAttributeValue<string>("ts_tablename")
                                    });

                                    //fileItem.TableRecordId = createdSharePointFile.ToString();
                                    fileItem.TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid");
                                }

                                //// Upload Attachment
                                //Task<bool> uploadTask = Task.Run(async () => await UploadAttachmentPowerAutomate(fileItem.FileId, fileItem.TableRecordId, fileItem.FileName));

                                //await Task.WhenAll(uploadTask);

                                //// If we have uploaded the file successfully, update the File
                                //if (uploadTask.Result)
                                //{
                                //    columns = new ColumnSet(true);

                                //    Guid fileGuid = new Guid(fileItem.FileId);

                                //    Entity fileRecord = svc.Retrieve("ts_file", fileGuid, columns);

                                //    fileRecord.Attributes["ts_uploadedtosharepoint"] = true;

                                //    svc.Update(fileRecord);
                                //}

                                counter++;
                                Console.WriteLine($"{counter} of {files.Entities.Count} completed");
                                Console.WriteLine($"---------------------------------------------");
                            }
                        }

                        foreach (var fileItem in fileItems.Where(x => x.UploadedToSharePoint == false && (x.ExemptionId != Guid.Empty || x.SecurityIncidentId != Guid.Empty)))
                        {
                            // Get the table name
                            string tableName = "";
                            string recordId = "";

                            if (fileItem.ExemptionId != Guid.Empty)
                            {
                                tableName = "Exemption";
                                recordId = fileItem.ExemptionId.ToString();
                            }

                            if (fileItem.SecurityIncidentId != Guid.Empty)
                            {
                                tableName = "Security Incident";
                                recordId = fileItem.SecurityIncidentId.ToString();
                            }

                            // Some Output
                            Console.WriteLine($"ID: {fileItem.FileId}");
                            Console.WriteLine($"File Name: {fileItem.FileName}");
                            Console.WriteLine($"{tableName}: {recordId}");

                            // Define the columns you want to retrieve (null retrieves all columns)
                            ColumnSet columns = new ColumnSet(true);

                            Entity sharePointFile = null;

                            EntityCollection sharePointFiles = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.Single_SharePointFile(recordId)));

                            sharePointFile = sharePointFiles.Entities.Count > 0 ? sharePointFiles.Entities[0] : null;

                            if (sharePointFile != null)
                            {
                                fileItem.TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid");

                                sharePointFileItems.Add(new SharePointFileItem
                                {
                                    SharePointFileId = sharePointFile.Id.ToString(),
                                    SharePointFileGroupId = sharePointFile.GetAttributeValue<EntityReference>("ts_sharepointfilegroup")?.Id.ToString(),
                                    TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid"),
                                    TableName = sharePointFile.GetAttributeValue<string>("ts_tablename")
                                });
                            }
                            else
                            {
                                // Create the SharePoint File
                                Entity newSharePointFile = new Entity("ts_sharepointfile");
                                newSharePointFile.Attributes["ts_tablename"] = tableName;
                                newSharePointFile.Attributes["ts_tablerecordid"] = recordId;

                                // Record the ID (GUID) of the new SharePoint File
                                Guid createdSharePointFile = svc.Create(newSharePointFile);

                                sharePointFiles = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.Single_SharePointFile(recordId)));

                                sharePointFile = sharePointFiles.Entities.Count > 0 ? sharePointFiles.Entities[0] : null;

                                // Add the SharePoint File to the list
                                sharePointFileItems.Add(new SharePointFileItem
                                {
                                    SharePointFileId = sharePointFile.Id.ToString(),
                                    SharePointFileGroupId = sharePointFile.GetAttributeValue<EntityReference>("ts_sharepointfilegroup")?.Id.ToString(),
                                    TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid"),
                                    TableName = sharePointFile.GetAttributeValue<string>("ts_tablename")
                                });

                                //fileItem.TableRecordId = createdSharePointFile.ToString();
                                fileItem.TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid");
                            }

                            counter++;
                            Console.WriteLine($"{counter} of {files.Entities.Count} completed");
                            Console.WriteLine($"---------------------------------------------");
                        }
                    }

                    // do the SharePoint File Group logic 
                    {
                        foreach (var sharePointFileItem in sharePointFileItems.OrderBy(x=> x.TableName))
                        {
                            if (sharePointFileItem.TableName == "Case" ||
                                sharePointFileItem.TableName == "Work Order" ||
                                sharePointFileItem.TableName == "Work Order Service Task")
                            {
                                // Check if it's a Case
                                if (sharePointFileItem.TableName == "Case")
                                {
                                    // Create the SharePoint File Group
                                    var sharePointFileGroup = GetOrCreateSharePointFileGroup(svc, sharePointFileItem);
                                }

                                // Check if it's a Work Order
                                if (sharePointFileItem.TableName == "Work Order")
                                {
                                    // Does the Work Order Have a Case?
                                    //var myWorkOrder = workOrderItems.FirstOrDefault(x => x.WorkOrderId == sharePointFileItem.TableRecordId);
                                    var myWorkOrder = GetRecordFromTable(svc, sharePointFileItem.TableRecordId, "msdyn_workorder");

                                    var sharePointFileGroup = new Entity();

                                    var caseId = myWorkOrder.GetAttributeValue<EntityReference>("msdyn_servicerequest");

                                    if (myWorkOrder!= null && caseId != null)
                                    {
                                        // Update the SharePointFile with the SharePointFileGroup of the Case
                                        //var caseSharePoint = sharePointFileItems.FirstOrDefault(x => x.TableRecordId == caseId.ToString());
                                        //NOTE: get the actual SharePoint file group from the database
                                        sharePointFileGroup = GetOrCreateSharePointFileGroup(svc, caseSharePoint);
                                    }
                                    else
                                    {
                                        // Give the Work Order it's own separate SharePointFileGroup
                                        sharePointFileGroup = GetOrCreateSharePointFileGroup(svc, sharePointFileItem);
                                    }
                                }

                            }
                            else
                            {
                                // Check if the SharePointFileGroup exists
                                if (string.IsNullOrWhiteSpace(sharePointFileItem.SharePointFileGroupId))
                                {
                                    var sharePointFileGroup = GetOrCreateSharePointFileGroup(svc, sharePointFileItem);
                                    sharePointFileItem.SharePointFileGroupId = sharePointFileGroup.Id.ToString();
                                }
                            }
                        }
                    }
                }
            }

            Console.WriteLine($"All Done!!!");
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

        private static async Task<bool> UploadAttachmentPowerAutomate(string fileId, string sharePointFileGroup, string fileName)
        {
            string flowEndpointUrl = "https://prod-08.canadacentral.logic.azure.com:443/workflows/5770469a718943aea1e4d87b9ec3c769/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=151Hbu33fqLvsCHsvQesCSMLHNreNezie9sOhAERvWE"; // Replace with your actual flow endpoint URL

            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri(flowEndpointUrl);
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, client.BaseAddress);

            // Serialize the byte array to a JSON string and convert it to Base64
            string jsonPayload = JsonConvert.SerializeObject(new { ts_fileid = fileId, ts_sharepointfileid = sharePointFileGroup, FileName= fileName });

            var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
            request.Content = content;
            var response = await MakeRequestAsync(request, client);
            Console.WriteLine(response);

            if (response == "File uploaded to SharePoint")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static async Task<string> MakeRequestAsync(HttpRequestMessage getRequest, HttpClient client)
        {
            var response = await client.SendAsync(getRequest).ConfigureAwait(false);
            var responseString = string.Empty;
            try
            {
                response.EnsureSuccessStatusCode();
                responseString = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            }
            catch (HttpRequestException)
            {
                // empty responseString
            }

            return responseString;
        }

        public static List<FileItem> GetFileItems(IOrganizationService svc, string fetchXml, string entityIdAttribute, string fileIdAttribute)
        {
            EntityCollection files = svc.RetrieveMultiple(new FetchExpression(fetchXml));
            List<FileItem> fileItems = new List<FileItem>();

            // Create a dictionary to map entityIdAttribute to corresponding property name
            var entityPropertyMap = new Dictionary<string, string>
            {
                { "incidentid", "CaseId" },
                { "msdyn_workorderid", "WorkOrderId" },
                { "msdyn_workorderservicetaskid", "WorkOrderServiceTaskId" },
                { "ovs_findingid", "FindingId" },
                { "msdyn_functionallocationid", "SiteId" },
                { "ovs_operationid", "OperationId" },
                { "accountid", "StakeholderId" }            
            };

            foreach (var file in files.Entities)
            {
                string fileId = file.GetAttributeValue<Guid>("ts_fileid").ToString();
                Guid recordId = file.GetAttributeValue<Guid>(entityIdAttribute);

                string propertyName = entityPropertyMap[entityIdAttribute];

                fileItems.Add(new FileItem
                {
                    FileId = fileId,
                    IsManyToMany = true,
                    FormIntegrationId = entityIdAttribute
                });

                // Use reflection to set the property dynamically based on entityIdAttribute
                var propertyInfo = typeof(FileItem).GetProperty(propertyName);

                if (propertyInfo != null)
                {
                    propertyInfo.SetValue(fileItems[fileItems.Count - 1], recordId);
                }
            }

            return fileItems;
        }

        public static Entity GetOrCreateSharePointFileGroup(IOrganizationService svc, SharePointFileItem sharePointFileItem)
        {
            Entity sharePointFileGroup = null;

            EntityCollection sharePointFileGroups = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.SharePointFileGroupBySharePointFile(sharePointFileItem.TableRecordId)));

            sharePointFileGroup = sharePointFileGroups.Entities.Count > 0 ? sharePointFileGroups.Entities[0] : null;

            if (sharePointFileGroup == null)
            {
                ColumnSet columns = new ColumnSet(true);

                Guid sharePointFileIdGuid = new Guid(sharePointFileItem.SharePointFileId);
                Entity sharePointFileRecord = svc.Retrieve("ts_sharepointfile", sharePointFileIdGuid, columns);

                if (sharePointFileRecord != null)
                {
                    Entity newSharePointFileGroup = new Entity("ts_sharepointfilegroup");
                    var newSharePointFileGroupID = svc.Create(newSharePointFileGroup);

                    sharePointFileRecord["ts_sharepointfilegroup"] = new EntityReference("ts_sharepointfilegroup", newSharePointFileGroupID);
                    svc.Update(sharePointFileRecord);

                    sharePointFileGroup = newSharePointFileGroup;
                }
            }

            return sharePointFileGroup;
        }

        public static Entity GetRecordFromTable(IOrganizationService svc, string recordId, string tableName)
        {
            Entity entity = null;
            ColumnSet columns = new ColumnSet(true);
            Guid recordIdGuid = new Guid(recordId);
            entity = svc.Retrieve(tableName, recordIdGuid, columns);
            return entity;
        }
    }

    public class FileItem
    {
        public string FileId { get; set; }
        public string FileName { get; set; }
        public bool? UploadedToSharePoint { get; set; }
        public Guid ExemptionId { get; set; }
        public Guid FindingId { get; set; }
        public Guid CaseId { get; set; }
        public Guid WorkOrderId { get; set; }
        public Guid OperationId { get; set; }
        public Guid SecurityIncidentId { get; set; }
        public Guid SiteId { get; set; }
        public Guid StakeholderId { get; set; }
        public Guid WorkOrderServiceTaskId { get; set; }
        public string TableRecordId { get; set; }
        public bool IsManyToMany { get; set; } = false;
        public string FormIntegrationId { get; set; }
        public List<FileItemGroup> FileItemGroups { get; set; } = new List<FileItemGroup>();
    }

    public class FileItemGroup
    {
        public Guid Id { get; set; }
        public string IdFieldName { get; set; }
    }

    public class SharePointFileItem
    {
        public string SharePointFileId { get; set; }
        public string SharePointFileGroupId { get; set; }
        public string TableRecordId { get; set; }
        public string TableName { get; set; }
    }

    public class Case
    {
        public string CaseId { get; set; }
        public string SharePointFileGroup { get; set; }
    }

    public class WorkOrder
    {
        public string WorkOrderId { get; set; }
        public string CaseId { get; set; }
        public string SharePointFileGroup { get; set; }
    }

    public class WorkOrderServiceTask
    {
        public string WorkOrderServiceTaskId { get; set; }
        public string WorkOrderId { get; set; }
    }
}
