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
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.SharePoint.Client.RecordsRepository;
using System.Text.RegularExpressions;

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
        static List<SecurityIncident> securityIncidents = new List<SecurityIncident>();
        static List<Exemption> exemptions = new List<Exemption>();

        public static string Case = "Case";
        public static string WorkOrder = "Work Order";
        public static string WorkOrderServiceTask = "Work Order Service Task";
        public static string Stakeholder = "Stakeholder";
        public static string Site = "Site";
        public static string Operation = "Operation";
        public static string SecurityIncident = "Security Incident";
        public static string Exemption = "Exemption";
        public static string CaseFr = "Cas";
        public static string WorkOrderFr = "Ordre de travail";
        public static string WorkOrderServiceTaskFr = "Tâche de service de l'ordre de travail";
        public static string StakeholderFr = "Partie prenante";
        public static string SiteFr = "Site";
        public static string OperationFr = "Opération";
        public static string SecurityIncidentFr = "Incidents de sûreté";
        public static string ExemptionFr = "Exemption";

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
                    //EntityCollection files = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.All_Files));
                    EntityCollection files = RetrieveAllRecordsUsingPaging(svc, new FetchExpression(FetchXMLExamples.All_Files));


                    // Get all Security Incidents
                    //EntityCollection securityIncidents = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.All_Security_Incidents));
                    EntityCollection securityIncidents = RetrieveAllRecordsUsingPaging(svc, new FetchExpression(FetchXMLExamples.All_Security_Incidents));

                    {
                        foreach (var securityIncident in securityIncidents.Entities)
                        {
                            string securityIncidentId = securityIncident.Id.ToString();
                            string securityIncidentName = securityIncident.GetAttributeValue<string> ("ts_name");

                            Program.securityIncidents.Add(new SecurityIncident
                            { 
                                SecurityIncidentId = securityIncidentId,
                                SecurityIncidentName = securityIncidentName
                            });
                        }
                    }

                    // Get all Exemptions
                    //EntityCollection exemptions = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.All_Exemptions));
                    EntityCollection exemptions = RetrieveAllRecordsUsingPaging(svc, new FetchExpression(FetchXMLExamples.All_Exemptions));

                    {
                        foreach (var exemption in exemptions.Entities)
                        {
                            string exemptionId = exemption.Id.ToString();
                            string exemptionNumber = exemption.GetAttributeValue<string>("ts_name");

                            Program.exemptions.Add(new Exemption
                            {
                                ExcemptionId = exemptionId,
                                ExemptionNumber = exemptionNumber
                            });
                        }
                    }

                    // get all individual file records
                    Console.WriteLine("Getting all individual file records.");
                    {
                        // use this to make sure we get random numbers every time
                        List<string> myRandomNumbers = new List<string>();
                        Random random = new Random();

                        // put all the files in a list
                        foreach (var file in files.Entities)
                        {

                            string randomNumber;

                            do
                            {
                                randomNumber = random.Next(10000, 99999).ToString();
                            } while (myRandomNumbers.Any(x => x == randomNumber));

                            myRandomNumbers.Add(randomNumber);

                            string fileId = file.GetAttributeValue<Guid>("ts_fileid").ToString();
                            string fileName = randomNumber + DateTime.UtcNow.ToString("yyyyMMddHHmmssffff") + " " + file.GetAttributeValue<String>("ts_file");
                            bool? uploadedToSharePoint = file.GetAttributeValue<bool>("ts_uploadedtosharepoint");
                            Guid exemptionId = file.GetAttributeValue<EntityReference>("ts_exemption")?.Id ?? Guid.Empty;
                            Guid securityIncidentId = file.GetAttributeValue<EntityReference>("ts_securityincident")?.Id ?? Guid.Empty;
                            string fileOwner = file.GetAttributeValue<AliasedValue>("FileOwner").Value.ToString();
                            string categoryEnglish = file.GetAttributeValue<AliasedValue>("CategoryEnglish")?.Value.ToString() ?? "";
                            string categoryFrench = file.GetAttributeValue<AliasedValue>("CategoryFrench")?.Value.ToString() ?? "";
                            string subCategoryEnglish = file.GetAttributeValue<AliasedValue>("SubCategoryEnglish")?.Value.ToString() ?? "Other";
                            subCategoryEnglish = string.IsNullOrWhiteSpace(subCategoryEnglish) ? "Other" : subCategoryEnglish;
                            string subCategoryFrench = file.GetAttributeValue<AliasedValue>("SubCategoryFrench")?.Value.ToString() ?? "Autre";
                            subCategoryFrench = string.IsNullOrWhiteSpace(subCategoryFrench) ? "Autre" : subCategoryFrench;
                            string fileDescription = file.GetAttributeValue<string>("ts_description")?.ToString() ?? "";
                            string formIntegrationID = "";
                            string tableRecordName = "";
                            string attachment = file.GetAttributeValue<Guid>("ts_attachment").ToString() ?? "";

                            if (exemptionId != Guid.Empty)
                            {
                                formIntegrationID = exemptionId.ToString();
                                tableRecordName = Program.exemptions.FirstOrDefault(x => x.ExcemptionId == formIntegrationID).ExemptionNumber;
                            }

                            if(securityIncidentId != Guid.Empty)
                            {
                                formIntegrationID = securityIncidentId.ToString();
                                tableRecordName = Program.securityIncidents.FirstOrDefault(x => x.SecurityIncidentId == formIntegrationID).SecurityIncidentName;
                            }

                            fileItems.Add(new FileItem
                            {
                                FileId = fileId,
                                FileName = fileName,
                                UploadedToSharePoint = uploadedToSharePoint,
                                ExemptionId = exemptionId,
                                SecurityIncidentId = securityIncidentId,
                                FormIntegrationId = formIntegrationID,
                                TableRecordName = tableRecordName,
                                FileOwner = fileOwner,
                                CategoryEnglish = categoryEnglish + " - " + subCategoryEnglish,
                                CategoryFrench = categoryFrench + " - " + subCategoryFrench,
                                FileDescription = fileDescription,
                                Attachment = attachment
                            });
                        }
                    }

                    // now get all the many-to-many records for file
                    Console.WriteLine("Getting all the many-to-many file records.");
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
                                        IdFieldName = itemId.FormIntegrationId,
                                        TableRecordName = itemId.TableRecordName
                                    });
                                }

                                if (itemId.WorkOrderId != Guid.Empty)
                                {
                                    fileItem.FileItemGroups.Add(new FileItemGroup
                                    {
                                        Id = itemId.WorkOrderId,
                                        IdFieldName = itemId.FormIntegrationId,
                                        TableRecordName = itemId.TableRecordName
                                    });
                                }

                                if (itemId.WorkOrderServiceTaskId != Guid.Empty)
                                {
                                    fileItem.FileItemGroups.Add(new FileItemGroup
                                    {
                                        Id = itemId.WorkOrderServiceTaskId,
                                        IdFieldName = itemId.FormIntegrationId,
                                        TableRecordName = itemId.TableRecordName
                                    });
                                }

                                if (itemId.StakeholderId != Guid.Empty)
                                {
                                    fileItem.FileItemGroups.Add(new FileItemGroup
                                    {
                                        Id = itemId.StakeholderId,
                                        IdFieldName = itemId.FormIntegrationId,
                                        TableRecordName = itemId.TableRecordName
                                    });
                                }

                                if (itemId.SiteId != Guid.Empty)
                                {
                                    fileItem.FileItemGroups.Add(new FileItemGroup
                                    {
                                        Id = itemId.SiteId,
                                        IdFieldName = itemId.FormIntegrationId,
                                        TableRecordName = itemId.TableRecordName
                                    });
                                }

                                if (itemId.OperationId != Guid.Empty)
                                {
                                    fileItem.FileItemGroups.Add(new FileItemGroup
                                    {
                                        Id = itemId.OperationId,
                                        IdFieldName = itemId.FormIntegrationId,
                                        TableRecordName = itemId.TableRecordName
                                    });
                                }
                            }
                        }
                    }

                    // do the SharePoint File logic 
                    {
                        // get all the ts_sharepointfile records that have already been created
                        EntityCollection sharePointFiles = RetrieveAllRecordsUsingPaging(svc, new FetchExpression(FetchXMLExamples.All_Existing_SharePointFiles));

                        counter = 0;
                        foreach (var fileItem in fileItems.Where(x => (x.UploadedToSharePoint == false || x.UploadedToSharePoint == null) && x.FileItemGroups.Count > 0))
                        {
                            foreach (var fileItemGroup in fileItem.FileItemGroups)
                            {
                                // Get the table name
                                string tableName = "";
                                string tableNameFr = "";

                                if (fileItemGroup.IdFieldName == "incidentid")
                                {
                                    tableName = Case;
                                    tableNameFr = CaseFr;
                                }

                                if (fileItemGroup.IdFieldName == "msdyn_workorderid")
                                {
                                    tableName = WorkOrder;
                                    tableNameFr = WorkOrderFr;
                                }

                                if (fileItemGroup.IdFieldName == "msdyn_workorderservicetaskid")
                                {
                                    tableName = WorkOrderServiceTask;
                                    tableNameFr = WorkOrderServiceTaskFr;
                                }

                                if (fileItemGroup.IdFieldName == "accountid")
                                {
                                    tableName = Stakeholder;
                                    tableNameFr = StakeholderFr;
                                }

                                if (fileItemGroup.IdFieldName == "msdyn_functionallocationid")
                                {
                                    tableName = Site;
                                    tableNameFr = SiteFr;
                                }

                                if (fileItemGroup.IdFieldName == "ovs_operationid")
                                {
                                    tableName = Operation;
                                    tableNameFr = OperationFr;
                                }

                                // Some Output
                                //Console.WriteLine($"ID: {fileItem.FileId}");
                                //Console.WriteLine($"File Name: {fileItem.FileName}");
                                //Console.WriteLine($"{tableName}: {fileItemGroup.Id}");

                                // Define the columns you want to retrieve (null retrieves all columns)
                                ColumnSet columns = new ColumnSet(true);

                                Entity sharePointFile = null;

                                // find out if the ts_sharepointfile already exists
                                sharePointFile = CheckIfSharePointFileExists(fileItemGroup, svc, sharePointFile, sharePointFiles,"");

                                if (sharePointFile != null)
                                {
                                    fileItem.TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid");

                                    fileItem.SharePointFileId = sharePointFile.Id.ToString();

                                    sharePointFileItems.Add(new SharePointFileItem
                                    {
                                        SharePointFileId = sharePointFile.Id.ToString(),
                                        SharePointFileGroupId = sharePointFile.GetAttributeValue<EntityReference>("ts_sharepointfilegroup")?.Id.ToString(),
                                        TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid"),
                                        TableName = sharePointFile.GetAttributeValue<string>("ts_tablename"),
                                        TableNameFrench = sharePointFile.GetAttributeValue<string>("ts_tablenamefrench"),
                                        TableRecordName = sharePointFile.GetAttributeValue<string>("ts_tablerecordname")
                                    });

                                    fileItem.SharePointTableName = sharePointFile.GetAttributeValue<string>("ts_tablename") + " - " + sharePointFile.GetAttributeValue<string>("ts_tablenamefrench");
                                    fileItem.SharePointTableRecordName = sharePointFile.GetAttributeValue<string>("ts_tablerecordname");
                                }
                                else
                                {
                                    bool fileOwnerValid = true;

                                    IsOwnerValid(fileItem, ref fileOwnerValid);

                                    if (fileOwnerValid)
                                    {
                                        // Create the SharePoint File
                                        Entity newSharePointFile = new Entity("ts_sharepointfile");
                                        newSharePointFile.Attributes["ts_tablename"] = tableName;
                                        newSharePointFile.Attributes["ts_tablenamefrench"] = tableNameFr;
                                        newSharePointFile.Attributes["ts_tablerecordid"] = fileItemGroup.Id.ToString();
                                        newSharePointFile.Attributes["ts_tablerecordname"] = fileItemGroup.TableRecordName;
                                        newSharePointFile.Attributes["ts_tablerecordowner"] = fileItem.FileOwner;

                                        // Record the ID (GUID) of the new SharePoint File
                                        Guid createdSharePointFile = svc.Create(newSharePointFile);

                                        sharePointFiles = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.Single_SharePointFile(fileItemGroup.Id.ToString())));

                                        sharePointFile = sharePointFiles.Entities.Count > 0 ? sharePointFiles.Entities[0] : null;

                                        fileItem.SharePointFileId = sharePointFile.Id.ToString();

                                        // Add the SharePoint File to the list
                                        sharePointFileItems.Add(new SharePointFileItem
                                        {
                                            SharePointFileId = sharePointFile.Id.ToString(),
                                            SharePointFileGroupId = sharePointFile.GetAttributeValue<EntityReference>("ts_sharepointfilegroup")?.Id.ToString(),
                                            TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid"),
                                            TableName = sharePointFile.GetAttributeValue<string>("ts_tablename"),
                                            TableNameFrench = sharePointFile.GetAttributeValue<string>("ts_tablenamefrench"),
                                            TableRecordName = sharePointFile.GetAttributeValue<string>("ts_tablerecordname")
                                        });

                                        //fileItem.TableRecordId = createdSharePointFile.ToString();
                                        fileItem.TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid");
                                        fileItem.SharePointTableName = sharePointFile.GetAttributeValue<string>("ts_tablename") + " - " + sharePointFile.GetAttributeValue<string>("ts_tablenamefrench");
                                        fileItem.SharePointTableRecordName = sharePointFile.GetAttributeValue<string>("ts_tablerecordname");
                                    }
                                }
                            }

                            counter++;
                            Console.WriteLine($"SharePoint File {counter} of {fileItems.Count(x => (x.UploadedToSharePoint == false || x.UploadedToSharePoint == null) && x.FileItemGroups.Count > 0)} completed");
                            Console.WriteLine($"---------------------------------------------");
                            //Console.Clear();
                        }

                        counter = 0;
                        foreach (var fileItem in fileItems.Where(x => x.UploadedToSharePoint == false && (x.ExemptionId != Guid.Empty || x.SecurityIncidentId != Guid.Empty)))
                        {
                            // Get the table name
                            string tableName = "";
                            string tableNameFr = "";
                            string recordId = "";

                            if (fileItem.ExemptionId != Guid.Empty)
                            {
                                tableName = Exemption;
                                tableNameFr = ExemptionFr;
                                recordId = fileItem.ExemptionId.ToString();
                            }

                            if (fileItem.SecurityIncidentId != Guid.Empty)
                            {
                                tableName = SecurityIncident;
                                tableNameFr = SecurityIncidentFr;
                                recordId = fileItem.SecurityIncidentId.ToString();
                            }

                            // Some Output
                            //Console.WriteLine($"ID: {fileItem.FileId}");
                            //Console.WriteLine($"File Name: {fileItem.FileName}");
                            //Console.WriteLine($"{tableName}: {recordId}");

                            // Define the columns you want to retrieve (null retrieves all columns)
                            ColumnSet columns = new ColumnSet(true);

                            Entity sharePointFile = null;

                            // find out if the ts_sharepointfile already exists
                            sharePointFile = CheckIfSharePointFileExists(null, svc, sharePointFile, sharePointFiles,recordId);

                            if (sharePointFile != null)
                            {
                                fileItem.TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid");

                                fileItem.SharePointFileId = sharePointFile.Id.ToString();

                                sharePointFileItems.Add(new SharePointFileItem
                                {
                                    SharePointFileId = sharePointFile.Id.ToString(),
                                    SharePointFileGroupId = sharePointFile.GetAttributeValue<EntityReference>("ts_sharepointfilegroup")?.Id.ToString(),
                                    TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid"),
                                    TableName = sharePointFile.GetAttributeValue<string>("ts_tablename"),
                                    TableNameFrench = sharePointFile.GetAttributeValue<string>("ts_tablenamefrench"),
                                    TableRecordName = sharePointFile.GetAttributeValue<string>("ts_tablerecordname")
                                });

                                fileItem.SharePointTableName = sharePointFile.GetAttributeValue<string>("ts_tablename") + " - " + sharePointFile.GetAttributeValue<string>("ts_tablenamefrench");
                                fileItem.SharePointTableRecordName = sharePointFile.GetAttributeValue<string>("ts_tablerecordname");

                            }
                            else
                            {
                                bool fileOwnerValid = true;
                                IsOwnerValid(fileItem, ref fileOwnerValid);

                                if (fileOwnerValid)
                                {
                                    // Create the SharePoint File
                                    Entity newSharePointFile = new Entity("ts_sharepointfile");
                                    newSharePointFile.Attributes["ts_tablename"] = tableName;
                                    newSharePointFile.Attributes["ts_tablenamefrench"] = tableNameFr;
                                    newSharePointFile.Attributes["ts_tablerecordid"] = recordId;
                                    newSharePointFile.Attributes["ts_tablerecordname"] = fileItem.TableRecordName;
                                    newSharePointFile.Attributes["ts_tablerecordowner"] = fileItem.FileOwner;

                                    // Record the ID (GUID) of the new SharePoint File
                                    Guid createdSharePointFile = svc.Create(newSharePointFile);

                                    sharePointFiles = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.Single_SharePointFile(recordId)));

                                    sharePointFile = sharePointFiles.Entities.Count > 0 ? sharePointFiles.Entities[0] : null;

                                    fileItem.SharePointFileId = sharePointFile.Id.ToString();

                                    // Add the SharePoint File to the list
                                    sharePointFileItems.Add(new SharePointFileItem
                                    {
                                        SharePointFileId = sharePointFile.Id.ToString(),
                                        SharePointFileGroupId = sharePointFile.GetAttributeValue<EntityReference>("ts_sharepointfilegroup")?.Id.ToString(),
                                        TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid"),
                                        TableName = sharePointFile.GetAttributeValue<string>("ts_tablename"),
                                        TableNameFrench = sharePointFile.GetAttributeValue<string>("ts_tablenamefrench"),
                                        TableRecordName = sharePointFile.GetAttributeValue<string>("ts_tablerecordname")
                                    });

                                    //fileItem.TableRecordId = createdSharePointFile.ToString();
                                    fileItem.TableRecordId = sharePointFile.GetAttributeValue<string>("ts_tablerecordid");

                                    fileItem.SharePointTableName = sharePointFile.GetAttributeValue<string>("ts_tablename") + " - " + sharePointFile.GetAttributeValue<string>("ts_tablenamefrench");
                                    fileItem.SharePointTableRecordName = sharePointFile.GetAttributeValue<string>("ts_tablerecordname");

                                }

                            }

                            counter++;
                            Console.WriteLine($"SharePoint File for Security Incident or Exemption {counter} of {fileItems.Count(x => x.UploadedToSharePoint == false && (x.ExemptionId != Guid.Empty || x.SecurityIncidentId != Guid.Empty))} completed");
                            Console.WriteLine($"---------------------------------------------");
                           // Console.Clear();
                        }
                    }

                    // do the SharePoint File Group logic 
                    //{
                    //    counter = 0;
                    //    foreach (var sharePointFileItem in sharePointFileItems.OrderBy(x=> x.TableName))
                    //    {
                    //        if (sharePointFileItem.TableName == Case ||
                    //            sharePointFileItem.TableName == WorkOrder ||
                    //            sharePointFileItem.TableName == WorkOrderServiceTask)
                    //        {
                    //            // Check if it's a Case
                    //            if (sharePointFileItem.TableName == Case)
                    //            {
                    //                // Create the SharePoint File Group
                    //                var sharePointFileGroup = GetOrCreateSharePointFileGroup(svc, sharePointFileItem);
                    //            }

                    //            // Check if it's a Work Order
                    //            if (sharePointFileItem.TableName == WorkOrder)
                    //            {
                    //                // Does the Work Order Have a Case?
                    //                var myWorkOrder = GetRecordFromTable(svc, sharePointFileItem.TableRecordId, "msdyn_workorder");

                    //                var caseSharePointFile = new Entity();

                    //                // Get the Case ID from the Work Order
                    //                var caseIdValue = myWorkOrder.GetAttributeValue<EntityReference>("msdyn_servicerequest");

                    //                if (myWorkOrder!= null && caseIdValue != null)
                    //                {
                    //                    string caseIdString = caseIdValue.Id.ToString();

                    //                    //Get the SharePointFile of the Case
                    //                    caseSharePointFile = GetSingleRecordFromTableFetchXML(svc, FetchXMLExamples.SharePointFileByTableRecordId(caseIdString));

                    //                    //If the SharePoint File doesn't exist for the Case, create it
                    //                    if (caseSharePointFile == null)
                    //                    {
                    //                        Entity newSharePointFile = new Entity("ts_sharepointfile");

                    //                        newSharePointFile.Attributes["ts_tablerecordid"] = caseIdString;
                    //                        newSharePointFile.Attributes["ts_tablename"] = Case;
                    //                        newSharePointFile.Attributes["ts_tablenamefrench"] = CaseFr;
                    //                        newSharePointFile.Attributes["ts_tablerecordname"] = caseIdValue.Name;
                    //                        newSharePointFile.Attributes["ts_tablerecordowner"] = sharePointFileItem.TableRecordOwner;

                    //                        Guid newSharePointFileID = svc.Create(newSharePointFile);

                    //                        var caseSharePointFileItem = new SharePointFileItem {
                    //                            SharePointFileId = newSharePointFileID.ToString(),
                    //                            TableRecordId = caseIdString,
                    //                            TableName = Case,
                    //                            TableNameFrench = CaseFr,
                    //                            TableRecordName = caseIdValue.Name
                    //                        };

                    //                        caseSharePointFile = GetOrCreateSharePointFileGroup(svc, caseSharePointFileItem,false);
                    //                    }

                    //                    string sharePointFileGroupId = "";

                    //                    // This is here because it can return a sharePointFile or sharePointFileGroup - sorry...
                    //                    if (caseSharePointFile.GetAttributeValue<EntityReference>("ts_sharepointfilegroup") == null)
                    //                    {
                    //                        sharePointFileGroupId = caseSharePointFile.Id.ToString();
                    //                    }
                    //                    else
                    //                    {
                    //                        sharePointFileGroupId = caseSharePointFile.GetAttributeValue<EntityReference>("ts_sharepointfilegroup").Id.ToString();
                    //                    }

                    //                    //Set the SharePointFileGroupId of the SharePointFile for the Work Order
                    //                    sharePointFileItem.SharePointFileGroupId = sharePointFileGroupId;

                    //                    // Update the SharePointFile
                    //                    GetOrCreateSharePointFileGroup(svc, sharePointFileItem,true);
                    //                }
                    //                else
                    //                {
                    //                    // Give the Work Order it's own separate SharePointFileGroup
                    //                    sharePointFileItem.SharePointFileGroupId = GetOrCreateSharePointFileGroup(svc, sharePointFileItem).Id.ToString();
                    //                }
                    //            }

                    //            // Check if it's a Work Order Service Task
                    //            if (sharePointFileItem.TableName == WorkOrderServiceTask)
                    //            {
                    //                // Does the Work Order Service Task Have a Work Order?
                    //                var myWorkOrderServiceTask = GetRecordFromTable(svc, sharePointFileItem.TableRecordId, "msdyn_workorderservicetask");

                    //                var workOrderSharePointFileGroup = new Entity();

                    //                var workOrderIdValue = myWorkOrderServiceTask.GetAttributeValue<EntityReference>("msdyn_workorder");

                    //                if (myWorkOrderServiceTask != null && workOrderIdValue != null)
                    //                {
                    //                    string workOrderIdString = workOrderIdValue.Id.ToString();

                    //                    //Get the SharePointFileGroup of the WorkOrder
                    //                    workOrderSharePointFileGroup = GetSingleRecordFromTableFetchXML(svc, FetchXMLExamples.SharePointFileByTableRecordId(workOrderIdString));

                    //                    //If the SharePoint File doesn't exist for the WorkOrder, create it
                    //                    if (workOrderSharePointFileGroup == null)
                    //                    {
                    //                        Entity newSharePointFile = new Entity("ts_sharepointfile");

                    //                        newSharePointFile.Attributes["ts_tablerecordid"] = workOrderIdString;
                    //                        newSharePointFile.Attributes["ts_tablename"] = WorkOrder;
                    //                        newSharePointFile.Attributes["ts_tablenamefrench"] = WorkOrderFr;
                    //                        newSharePointFile.Attributes["ts_tablerecordname"] = workOrderIdValue.Name;
                    //                        newSharePointFile.Attributes["ts_tablerecordowner"] = sharePointFileItem.TableRecordOwner;

                    //                        Guid newSharePointFileID = svc.Create(newSharePointFile);

                    //                        var workOrderSharePointFileItem = new SharePointFileItem
                    //                        {
                    //                            SharePointFileId = newSharePointFileID.ToString(),
                    //                            TableRecordId = workOrderIdString,
                    //                            TableName = Case,
                    //                            TableNameFrench = CaseFr,
                    //                            TableRecordName = workOrderIdValue.Name
                    //                        };

                    //                        workOrderSharePointFileGroup = GetOrCreateSharePointFileGroup(svc, workOrderSharePointFileItem, false);
                    //                    }

                    //                    string sharePointFileGroupId = "";

                    //                    // This is here because it can return a sharePointFile or sharePointFileGroup - sorry...
                    //                    if (workOrderSharePointFileGroup.GetAttributeValue<EntityReference>("ts_sharepointfilegroup") == null)
                    //                    {
                    //                        sharePointFileGroupId = workOrderSharePointFileGroup.Id.ToString();
                    //                    }
                    //                    else
                    //                    {
                    //                        sharePointFileGroupId = workOrderSharePointFileGroup.GetAttributeValue<EntityReference>("ts_sharepointfilegroup").Id.ToString();
                    //                    }

                    //                    //Set the SharePointFileGroupId of the SharePointFile for the Work Order
                    //                    sharePointFileItem.SharePointFileGroupId = sharePointFileGroupId;

                    //                    // Update the SharePointFile
                    //                    GetOrCreateSharePointFileGroup(svc, sharePointFileItem, true);
                    //                }
                    //                else
                    //                {
                    //                    // Give the Work Order Service Task it's own separate SharePointFileGroup
                    //                    sharePointFileItem.SharePointFileGroupId = GetOrCreateSharePointFileGroup(svc, sharePointFileItem).Id.ToString();
                    //                }
                    //            }
                    //        }
                    //        else
                    //        {
                    //            // Check if the SharePointFileGroup exists
                    //            if (string.IsNullOrWhiteSpace(sharePointFileItem.SharePointFileGroupId))
                    //            {
                    //                var sharePointFileGroup = GetOrCreateSharePointFileGroup(svc, sharePointFileItem);
                    //                sharePointFileItem.SharePointFileGroupId = sharePointFileGroup.Id.ToString();
                    //            }
                    //        }

                    //        counter++;
                    //        Console.WriteLine($"SharePoint File Group {counter} of {sharePointFileItems.Count()} completed");
                    //        Console.WriteLine($"---------------------------------------------");
                    //    }
                    //}

                    // Upload all the files to SharePoint
                    {
                        int fileCounter = 0;
                        int totalFileCount = fileItems.Count(x => x.UploadedToSharePoint == false && x.SharePointFileId != null && x.Attachment != "00000000-0000-0000-0000-000000000000");

                        foreach (var fileItem in fileItems.Where(x => x.UploadedToSharePoint == false &&
                        x.SharePointFileId != null &&
                        x.Attachment != "00000000-0000-0000-0000-000000000000"))
                        {
                            bool fileOwnerValid = true;

                            IsOwnerValid(fileItem, ref fileOwnerValid);

                            IsFileCategoryValid(fileItem, fileItem.FileOwner);

                            if (fileOwnerValid)
                            {
                                // Upload Attachment
                                Task<bool> uploadTask = UploadAttachmentPowerAutomateAsync(fileItem, fileCounter, totalFileCount);

                                await Task.WhenAll(uploadTask);

                                // If we have uploaded the file successfully, update the File
                                if (uploadTask.Result)
                                {
                                    ColumnSet columns = new ColumnSet(true);

                                    Guid fileGuid = new Guid(fileItem.FileId);

                                    Entity fileRecord = svc.Retrieve("ts_file", fileGuid, columns);

                                    fileRecord.Attributes["ts_uploadedtosharepoint"] = true;

                                    svc.Update(fileRecord);
                                }

                                fileCounter++;
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

        private static async Task<bool> UploadAttachmentPowerAutomateAsync(FileItem fileItem, int count, int totalCount)
        {
            // Move Files to SharePoint - ROMTS-GSRST.Flows
            //string devURL = "https://prod-08.canadacentral.logic.azure.com:443/workflows/5770469a718943aea1e4d87b9ec3c769/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=151Hbu33fqLvsCHsvQesCSMLHNreNezie9sOhAERvWE";

            string prodURL = "https://prod-08.canadacentral.logic.azure.com:443/workflows/3578d30a6d1f481a9b344afea45b3497/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=gQNND8iNEBLaBI0uF2vfWuw1jd1fsicMwyyS1Nae0kQ";

            //string flowEndpointUrl = devURL;
            string flowEndpointUrl = prodURL;

            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri(flowEndpointUrl);
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, client.BaseAddress);

            // Remove any invalid characters in the file name
            string fileName = fileItem.FileName;
            string extension = System.IO.Path.GetExtension(fileName);

            // List of invalid characters
            List<char> invalidChars = new List<char> { '~', '#', '%', '&', '*', '{', '}', '\\', ':', '<', '>', '?', '/', '+', '|', '"' };

            foreach (char c in invalidChars)
            {
                fileName = fileName.Replace(c.ToString(), string.Empty);
            }

            // Ensure that the file extension is preserved
            if (!fileName.EndsWith(extension))
            {
                fileName += extension;
            }

            // Apply the changes to fileItem.FileName
            fileItem.FileName = fileName;

            // Serialize the byte array to a JSON string and convert it to Base64
            string jsonPayload = JsonConvert.SerializeObject(new { 
                ts_fileid = fileItem.FileId, 
                ts_sharepointfileid = fileItem.SharePointFileId, 
                FileName = fileItem.FileName,
                CategoryEnglish = fileItem.CategoryEnglish,
                CategoryFrench = fileItem.CategoryFrench,
                FileDescription = fileItem.FileDescription,
                FileOwner = fileItem.FileOwner,
                TableName = fileItem.SharePointTableName,
                TableRecordName = fileItem.SharePointTableRecordName
            });

            var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
            request.Content = content;
            var response = await MakeRequestAsync(request, client);

            if (response == "File uploaded to SharePoint")
            {
                Console.WriteLine($"Number( {count} of {totalCount} ) File: {fileItem.FileName} -  {response}");
                return true;
            }
            else
            {
                Console.WriteLine($"Number ( {count} of {totalCount} ) File: {fileItem.FileName} -  Was not uploaded due to an error");
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
            //EntityCollection files = svc.RetrieveMultiple(new FetchExpression(fetchXml));

            EntityCollection files = RetrieveAllRecordsUsingPaging(svc, new FetchExpression(fetchXml));

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
                { "accountid", "StakeholderId" },
                { "tablerecordname", "TableRecordName" }
            };

            foreach (var file in files.Entities)
            {
                string fileId = file.GetAttributeValue<Guid>("ts_fileid").ToString();
                Guid recordId = file.GetAttributeValue<Guid>(entityIdAttribute);
                string tableRecordName = "";

                // in case we have Stakeholders that have only the Account Named filled out and not the Legal Name
                if (file.GetAttributeValue<AliasedValue>("tablerecordname") is null && file.LogicalName == "ts_files_accounts")
                {
                    tableRecordName = file.GetAttributeValue<AliasedValue>("tablerecordnamebackup").Value?.ToString();

                    if (tableRecordName.Contains("::"))
                    {
                        string[] parts = tableRecordName.Split(new string[] { "::" }, StringSplitOptions.None);
                        tableRecordName = parts[0];
                    }

                    if (tableRecordName.Length > 100)
                    {
                        tableRecordName = tableRecordName.Substring(0, 100);
                    }
                }
                else
                {
                    tableRecordName = file.GetAttributeValue<AliasedValue>("tablerecordname").Value?.ToString();

                    if (tableRecordName.Contains("::"))
                    {
                        string[] parts = tableRecordName.Split(new string[] { "::" }, StringSplitOptions.None);
                        tableRecordName = parts[0];
                    }

                    if (tableRecordName.Length > 100)
                    {
                        tableRecordName = tableRecordName.Substring(0, 100);
                    }
                }


                string propertyName = entityPropertyMap[entityIdAttribute];

                fileItems.Add(new FileItem
                {
                    FileId = fileId,
                    IsManyToMany = true,
                    FormIntegrationId = entityIdAttribute,
                    TableRecordName =  tableRecordName
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

        public static Entity GetOrCreateSharePointFileGroup(IOrganizationService svc, SharePointFileItem sharePointFileItem, bool usingExistingSharePointFileGroup = false)
        {
            // This will get or create the SharePoint File Group and update the SharePoint File
            Entity sharePointFileGroup = null;

            EntityCollection sharePointFiles = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.SharePointFileByTableRecordId(sharePointFileItem.TableRecordId)));

            sharePointFileGroup = sharePointFiles.Entities.Count > 0 ? sharePointFiles.Entities[0] : null;

            if (sharePointFileGroup == null)
            {
                ColumnSet columns = new ColumnSet(true);

                Guid sharePointFileIdGuid = new Guid(sharePointFileItem.SharePointFileId);
                Entity sharePointFileRecord = svc.Retrieve("ts_sharepointfile", sharePointFileIdGuid, columns);

                if (sharePointFileRecord != null)
                {
                    Entity newSharePointFileGroup = new Entity("ts_sharepointfilegroup");

                    Guid sharePointFileGroupID = Guid.Empty;

                    if (!usingExistingSharePointFileGroup)
                    {
                        sharePointFileGroupID = svc.Create(newSharePointFileGroup);
                        
                        // prevent an error by getting the new SharePoint File Group
                        newSharePointFileGroup = svc.Retrieve("ts_sharepointfilegroup", sharePointFileGroupID, columns);
                    }
                    else
                    {
                        sharePointFileGroupID = new Guid(sharePointFileItem.SharePointFileGroupId);
                    }

                    sharePointFileRecord["ts_sharepointfilegroup"] = new EntityReference("ts_sharepointfilegroup", sharePointFileGroupID);
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

        public static Entity GetSingleRecordFromTableFetchXML(IOrganizationService svc, string fetchXML)
        {
            Entity entity = null;

            EntityCollection entityCollection = svc.RetrieveMultiple(new FetchExpression(fetchXML));

            entity = entityCollection.Entities.Count > 0 ? entityCollection.Entities[0] : null;

            return entity;
        }

        public static void IsOwnerValid(FileItem fileItem, ref bool fileOwnerValid)
        {
            if (fileItem.FileOwner.Contains("Aviation Security"))
            {
                fileItem.FileOwner = "Aviation Security";
            }
            else if (fileItem.FileOwner.Contains("Intermodal Surface Security Oversight (ISSO)"))
            {
                fileItem.FileOwner = "Intermodal Surface Security Oversight (ISSO)";
            }
            else
            {
                fileOwnerValid = false;
            }
        }

        public static void IsFileCategoryValid(FileItem fileItem,string owner)
        {
            List<AvSecFileCategory> avsecFileCategories = new List<AvSecFileCategory>
            {
                new AvSecFileCategory { EnglishName = "Supporting Documentation - Email communication", FrenchName = "Document(s) justificatifs - Correspondance par courriel" },
                new AvSecFileCategory { EnglishName = "Supporting Documentation - Photograph", FrenchName = "Document(s) justificatifs - Photographie" },
                new AvSecFileCategory { EnglishName = "Supporting Documentation - Site map / diagram /schematic", FrenchName = "Document(s) justificatifs - Plan du site / diagramme / schéma" },
                new AvSecFileCategory { EnglishName = "Supporting Documentation - Letter", FrenchName = "Document(s) justificatifs - Lettre" },
                new AvSecFileCategory { EnglishName = "Supporting Documentation - Training Record", FrenchName = "Document(s) justificatifs - Dossier de formation" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Emergency Response Plan (ERP)", FrenchName = "Document(s) des intervenants - Plan d'intervention d'urgence" },
                new AvSecFileCategory { EnglishName = "Supporting Documentation - Note to file", FrenchName = "Document(s) justificatifs - Note au dossier" },
                new AvSecFileCategory { EnglishName = "Supporting Documentation - Inspection Questionnaire", FrenchName = "Document(s) justificatifs - Questionnaire d'inspection" },
                new AvSecFileCategory { EnglishName = "Supporting Documentation - Incident Report", FrenchName = "Document(s) justificatifs - Rapport d'incident" },
                new AvSecFileCategory { EnglishName = "Supporting Documentation - Other", FrenchName = "Document(s) justificatifs - Autre" },
                new AvSecFileCategory { EnglishName = "Internal Reference Material - Standard Operation Procedure (SOP)", FrenchName = "Élément(s) de référence interne - Procédure opérationnelle normalisée (PON)" },
                new AvSecFileCategory { EnglishName = "Internal Reference Material - User Guides", FrenchName = "Élément(s) de référence interne - Guide de l'utilisateur" },
                new AvSecFileCategory { EnglishName = "Internal Reference Material - Staff Instructions (SI)", FrenchName = "Élément(s) de référence interne - Instruction spéciale (IS)" },
                new AvSecFileCategory { EnglishName = "Internal Reference Material - Legislation", FrenchName = "Élément(s) de référence interne - Législation" },
                new AvSecFileCategory { EnglishName = "Internal Reference Material - Exemption", FrenchName = "Élément(s) de référence interne - Exemption" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Stakeholder Contact Information", FrenchName = "Document(s) des intervenants - Information sur les personnes-ressources des intervenants ou des partenaires" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Security Plan (SP)", FrenchName = "Document(s) des intervenants - Plan de sûreté (PS)" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Risk Assessment (SRA)", FrenchName = "Document(s) des intervenants - Évaluation des risques" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Railway Carrier Profile", FrenchName = "Document(s) des intervenants - Profil du transporteur ferroviaire" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Railway Loader Profile", FrenchName = "Document(s) des intervenants - Profil du chargeur ferroviaire" },
                new AvSecFileCategory { EnglishName = "Supporting Documentation - Training Material", FrenchName = "Document(s) justificatifs - Matériel de formation" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Canadian Aviation Document (CAD)", FrenchName = "Document(s) des intervenants - Document d'aviation Canadien (DAC)" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Railway Operating Certificate (ROC)", FrenchName = "Document(s) des intervenants - Certificat d'exploitation de chemin de fer" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Exercise", FrenchName = "Document(s) des intervenants - Exercice" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Other", FrenchName = "Document(s) des intervenants - Autre" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Site map / diagram / schematic", FrenchName = "Document(s) des intervenants - Plan du site / diagramme / schéma" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Training Record", FrenchName = "Document(s) des intervenants - Dossier de formation" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Passenger Profile", FrenchName = "Document(s) des intervenants - Profil de compagnie de transport de voyageurs" },
                new AvSecFileCategory { EnglishName = "Stakeholder Documentation - Training Material", FrenchName = "Document(s) des intervenants - Matériel de formation" },
                new AvSecFileCategory { EnglishName = "Supporting Documentation - Findings Report", FrenchName = "Document(s) justificatifs - Rapports de constatations" }
            };

            List<IssoFileCategory> issoCategories = new List<IssoFileCategory>
            {
                new IssoFileCategory { EnglishName = "Supporting Documentation - Email communication", FrenchName = "Document(s) justificatifs - Correspondance par courriel" },
                new IssoFileCategory { EnglishName = "Supporting Documentation - Photograph", FrenchName = "Document(s) justificatifs - Photographie" },
                new IssoFileCategory { EnglishName = "Supporting Documentation - Site map / diagram /schematic", FrenchName = "Document(s) justificatifs - Plan du site / diagramme / schéma" },
                new IssoFileCategory { EnglishName = "Supporting Documentation - Letter", FrenchName = "Document(s) justificatifs - Lettre" },
                new IssoFileCategory { EnglishName = "Supporting Documentation - Training Record", FrenchName = "Document(s) justificatifs - Dossier de formation" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Emergency Response Plan (ERP)", FrenchName = "Document(s) des intervenants - Plan d'intervention d'urgence" },
                new IssoFileCategory { EnglishName = "Supporting Documentation - Note to file", FrenchName = "Document(s) justificatifs - Note au dossier" },
                new IssoFileCategory { EnglishName = "Supporting Documentation - Inspection Questionnaire", FrenchName = "Document(s) justificatifs - Questionnaire d'inspection" },
                new IssoFileCategory { EnglishName = "Supporting Documentation - Incident Report", FrenchName = "Document(s) justificatifs - Rapport d'incident" },
                new IssoFileCategory { EnglishName = "Supporting Documentation - Other", FrenchName = "Document(s) justificatifs - Autre" },
                new IssoFileCategory { EnglishName = "Internal Reference Material - Standard Operation Procedure (SOP)", FrenchName = "Élément(s) de référence interne - Procédure opérationnelle normalisée (PON)" },
                new IssoFileCategory { EnglishName = "Internal Reference Material - User Guides", FrenchName = "Élément(s) de référence interne - Guide de l'utilisateur" },
                new IssoFileCategory { EnglishName = "Internal Reference Material - Staff Instructions (SI)", FrenchName = "Élément(s) de référence interne - Instruction spéciale (IS)" },
                new IssoFileCategory { EnglishName = "Internal Reference Material - Legislation", FrenchName = "Élément(s) de référence interne - Législation" },
                new IssoFileCategory { EnglishName = "Internal Reference Material - Exemption", FrenchName = "Élément(s) de référence interne - Exemption" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Stakeholder Contact Information", FrenchName = "Document(s) des intervenants - Information sur les personnes-ressources des intervenants ou des partenaires" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Security Plan (SP)", FrenchName = "Document(s) des intervenants - Plan de sûreté (PS)" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Risk Assessment (SRA)", FrenchName = "Document(s) des intervenants - Évaluation des risques" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Railway Carrier Profile", FrenchName = "Document(s) des intervenants - Profil du transporteur ferroviaire" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Railway Loader Profile", FrenchName = "Document(s) des intervenants - Profil du chargeur ferroviaire" },
                new IssoFileCategory { EnglishName = "Supporting Documentation - Training Material", FrenchName = "Document(s) justificatifs - Matériel de formation" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Canadian Aviation Document (CAD)", FrenchName = "Document(s) des intervenants - Document d'aviation Canadien (DAC)" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Railway Operating Certificate (ROC)", FrenchName = "Document(s) des intervenants - Certificat d'exploitation de chemin de fer" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Exercise", FrenchName = "Document(s) des intervenants - Exercice" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Other", FrenchName = "Document(s) des intervenants - Autre" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Site map / diagram / schematic", FrenchName = "Document(s) des intervenants - Plan du site / diagramme / schéma" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Training Record", FrenchName = "Document(s) des intervenants - Dossier de formation" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Passenger Profile", FrenchName = "Document(s) des intervenants - Profil de compagnie de transport de voyageurs" },
                new IssoFileCategory { EnglishName = "Stakeholder Documentation - Training Material", FrenchName = "Document(s) des intervenants - Matériel de formation" },
                new IssoFileCategory { EnglishName = "Supporting Documentation - Findings Report", FrenchName = "Document(s) justificatifs - Rapports de constatations" },
                new IssoFileCategory { EnglishName = "Internal Reference material - Other", FrenchName = "Élément(s) de référence interne - Autre" }
            };

            string myFileCategory = fileItem.CategoryEnglish;

            if (owner == "Aviation Security")
            {
                if (avsecFileCategories.Any(x => x.EnglishName == myFileCategory))
                {
                    // do nothing since the category is good
                }
                else
                {
                    fileItem.CategoryEnglish = "Supporting Documentation - Other";
                    fileItem.CategoryFrench = "Document(s) justificatifs - Autre";
                }
            }
            else if(owner == "Intermodal Surface Security Oversight (ISSO)")
            {
                if (issoCategories.Any(x => x.EnglishName == myFileCategory))
                {
                    // do nothing since the category is good
                }
                else
                {
                    fileItem.CategoryEnglish = "Supporting Documentation - Other";
                    fileItem.CategoryFrench = "Document(s) justificatifs - Autre";
                }
            }
        }

        public static EntityCollection RetrieveAllRecordsUsingPaging(IOrganizationService service, FetchExpression fetchExpression)
        {
            // Use the FetchXmlToQueryExpressionRequest message to convert the FetchExpression to a QueryExpression.
            FetchXmlToQueryExpressionRequest conversionRequest = new FetchXmlToQueryExpressionRequest
            {
                FetchXml = fetchExpression.Query
            };

            FetchXmlToQueryExpressionResponse conversionResponse = (FetchXmlToQueryExpressionResponse)service.Execute(conversionRequest);

            // The QueryExpression is now available.
            QueryExpression query = conversionResponse.Query;

            var pageNumber = 1;
            var pagingCookie = string.Empty;
            var result = new EntityCollection();
            EntityCollection resp = null;

            do
            {
                if (pageNumber != 1)
                {
                    query.PageInfo.PageNumber = pageNumber;
                    query.PageInfo.PagingCookie = pagingCookie;
                }

                resp = service.RetrieveMultiple(query);

                if (resp.MoreRecords)
                {
                    pageNumber++;
                    pagingCookie = resp.PagingCookie;
                }

                // Add the result from RetrieveMultiple to the EntityCollection to be returned.
                result.Entities.AddRange(resp.Entities);
            }
            while (resp != null && resp.MoreRecords);

            return result;
        }
    
        public static Entity CheckIfSharePointFileExists(FileItemGroup fileItemGroup, CrmServiceClient svc,Entity sharePointFile,EntityCollection sharePointFiles, string recordId)
        {
            // find out if the ts_sharepointfile already exists
            if (sharePointFiles.Entities.Count > 0)
            {
                //sharePointFile = sharePointFiles.Entities.FirstOrDefault(e => e.Attributes["ts_tablerecordid"].ToString().ToUpper() == fileItem.TableRecordId.ToUpper());
                foreach (var item in sharePointFiles.Entities)
                {
                    if (fileItemGroup != null)
                    {
                        if (item.GetAttributeValue<string>("ts_tablerecordid").ToUpper() == fileItemGroup.Id.ToString().ToUpper())
                        {
                            sharePointFile = item;
                            break;
                        }
                    }
                }
            }

            // if it is null - then do a check just to be sure
            if (sharePointFile == null)
            {
                if (fileItemGroup != null)
                {
                    sharePointFiles = svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.Single_SharePointFile(fileItemGroup.Id.ToString())));

                }
                else
                {
                    svc.RetrieveMultiple(new FetchExpression(FetchXMLExamples.Single_SharePointFile(recordId)));
                }

                sharePointFile = sharePointFiles.Entities.Count > 0 ? sharePointFiles.Entities[0] : null;
            }

            return sharePointFile;
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
        public string TableRecordName { get; set; }
        public string TableRecordId { get; set; }
        public bool IsManyToMany { get; set; } = false;
        public string FormIntegrationId { get; set; }
        public List<FileItemGroup> FileItemGroups { get; set; } = new List<FileItemGroup>();
        public string FileOwner { get; set; }
        public string CategoryEnglish { get; set; }
        public string CategoryFrench { get; set; }
        public string FileDescription { get; set; }
        public string SharePointFileId { get; set; }
        public string Attachment { get; set; }
        public string SharePointTableName { get; set; }
        public string SharePointTableRecordName { get; set; }
    }

    public class FileItemGroup
    {
        public Guid Id { get; set; }
        public string IdFieldName { get; set; }
        public string TableRecordName { get; set; }
    }

    public class SharePointFileItem
    {
        public string SharePointFileId { get; set; }
        public string SharePointFileGroupId { get; set; }
        public string TableRecordId { get; set; }
        public string TableName { get; set; }
        public string TableNameFrench { get; set; }
        public string TableRecordName { get; set; }
        public string TableRecordOwner { get; set; }
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

    public class SecurityIncident
    {
        public string SecurityIncidentId { get; set; }
        public string SecurityIncidentName { get; set; }
    }

    public class Exemption
    {
        public string ExcemptionId { get; set; }
        public string ExemptionNumber { get; set; }
    }

    public class AvSecFileCategory
    {
        public string EnglishName { get; set; }
        public string FrenchName { get; set; }
    }

    public class IssoFileCategory
    {
        public string EnglishName { get; set; }
        public string FrenchName { get; set; }
    }

}
