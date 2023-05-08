using Microsoft.AspNetCore.StaticFiles;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Extensions.Configuration;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.PowerPlatform.Dataverse.Client.Extensions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

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
                // Lookups
                {
                    // Retrieve all Security Incident Types
                    {
                        // English
                        EntityCollection securityIncidentTypesEnglish = svc.RetrieveMultiple(new FetchExpression(LookupFetchXML.SecurityIncidents()));

                        // French - Use this to retrieve the French names
                        EntityCollection securityIncidentTypesFrench = svc.RetrieveMultiple(new FetchExpression(LookupFetchXML.SecurityIncidents(false)));

                        //Console.WriteLine("Security Incident Types:");

                        //foreach (var securityIncident in securityIncidentTypesEnglish.Entities)
                        //{
                        //    Console.WriteLine($"ID: {securityIncident.Id.ToString()}");
                        //    Console.WriteLine($"Name: {securityIncident.GetAttributeValue<String>(LookupColumns.SecurityIncidentType.EnglishName)}");
                        //    Console.WriteLine("");
                        //}

                        //Console.WriteLine("End of Security Incident Types-------------------");
                    }

                    // Retrieve all Target Elements
                    {
                        // English
                        EntityCollection targetElementEnglish = svc.RetrieveMultiple(new FetchExpression(LookupFetchXML.TargetElements()));

                        // French - Use this to retrieve the French names
                        EntityCollection targetElementFrench = svc.RetrieveMultiple(new FetchExpression(LookupFetchXML.TargetElements(false)));

                        //Console.WriteLine("Target Elements:");

                        //foreach (var targetElement in targetElementEnglish.Entities)
                        //{
                        //    Console.WriteLine($"ID: {targetElement.Id.ToString()}");
                        //    Console.WriteLine($"Name: {targetElement.GetAttributeValue<String>(LookupColumns.TargetElement.EnglishName)}");
                        //    Console.WriteLine("");
                        //}

                        //Console.WriteLine("End of Target Elements-------------------");
                    }

                    // Retrieve all Reporting Companies
                    // Reporting Company and Stakeholder reside in the same table in ROM
                    {
                        // English
                        EntityCollection reportingCompany = svc.RetrieveMultiple(new FetchExpression(LookupFetchXML.ReportingCompany()));

                        //Console.WriteLine("Reporting Companies:");

                        //foreach (var reportingcompany in reportingCompany.Entities)
                        //{
                        //    Console.WriteLine($"ID: {reportingcompany.Id.ToString()}");
                        //    Console.WriteLine($"Name: {reportingcompany.GetAttributeValue<String>(LookupColumns.ReportingCompany.Name)}");
                        //    Console.WriteLine("");
                        //}

                        //Console.WriteLine("End of Reporting Companies-------------------");
                    }

                    // Retrieve all Stakeholder Operation Types
                    {
                        // English
                        EntityCollection stakeholderOperationTypesEnglish = svc.RetrieveMultiple(new FetchExpression(LookupFetchXML.StakeholderOperationType()));

                        // French - Use this to retrieve the French names
                        EntityCollection stakeholderOperationTypesFrench = svc.RetrieveMultiple(new FetchExpression(LookupFetchXML.TargetElements(false)));

                        //Console.WriteLine("Stakeholder Operation Types:");

                        //foreach (var stakeholderOperationType in stakeholderOperationTypesEnglish.Entities)
                        //{
                        //    Console.WriteLine($"ID: {stakeholderOperationType.Id.ToString()}");
                        //    Console.WriteLine($"Name: {stakeholderOperationType.GetAttributeValue<String>(LookupColumns.StakeholderOperationType.EnglishName)}");
                        //    Console.WriteLine("");
                        //}

                        //Console.WriteLine("End of Stakeholder Operation Types-------------------");
                    }

                    // Retrieve all Regions
                    {
                        // English
                        EntityCollection regionsEnglish = svc.RetrieveMultiple(new FetchExpression(LookupFetchXML.Region()));

                        // French - Use this to retrieve the French names
                        EntityCollection regionsFrench = svc.RetrieveMultiple(new FetchExpression(LookupFetchXML.Region(false)));

                        //Console.WriteLine("Regions:");

                        //foreach (var region in regionsEnglish.Entities)
                        //{
                        //    Console.WriteLine($"ID: {region.Id.ToString()}");
                        //    Console.WriteLine($"Name: {region.GetAttributeValue<String>(LookupColumns.Region.EnglishName)}");
                        //    Console.WriteLine("");
                        //}

                        //Console.WriteLine("End of Regions-------------------");
                    }

                    // Retrieve all Sites
                    // Site, SubSite, Origin, and Destination reside in the same table in ROM
                    {
                        // English
                        EntityCollection sites = svc.RetrieveMultiple(new FetchExpression(LookupFetchXML.Site()));

                        //Console.WriteLine("Sites:");

                        //foreach (var site in sites.Entities)
                        //{
                        //    Console.WriteLine($"ID: {site.Id.ToString()}");
                        //    Console.WriteLine($"Name: {site.GetAttributeValue<String>(LookupColumns.Site.Name)}");
                        //    Console.WriteLine($"Province: {site.GetAttributeValue<String>(LookupColumns.Site.Province)}");
                        //    Console.WriteLine($"Longitude: {site.GetAttributeValue<double>(LookupColumns.Site.Longitude).ToString()}");
                        //    Console.WriteLine($"Latitude: {site.GetAttributeValue<double>(LookupColumns.Site.Latitude).ToString()}");
                        //    Console.WriteLine("");
                        //}

                        //Console.WriteLine("End of Sites-------------------");
                    }
                }

                // Choices
                {
                    // Retrieve all Programs
                    {
                        var programs = Choice.GetChoices(svc.GetGlobalOptionSetMetadata(ChoiceOptions.Program.GlobalOptionSetName));

                        //Console.WriteLine("Programs:");

                        //foreach (var program in programs)
                        //{
                        //    Console.WriteLine($"Value: {program.Value}");
                        //    Console.WriteLine($"English Name: {program.EnglishLabel}");
                        //    Console.WriteLine($"French Name: {program.FrenchLabel}");
                        //    Console.WriteLine("");
                        //}

                        //Console.WriteLine("End of Programs-------------------");
                    }
                }

                //Console.WriteLine("- End Of Program -");

                // EXAMPLE - Create a Security Incident with a file attachment
                {
                    Entity newSecurityIncident = new Entity("ts_securityincident");

                    // Get a mode
                    var selectedMode = 717750002;

                    // Set the mode for the new Security Incident
                    newSecurityIncident.Attributes["ts_mode"] = new OptionSetValue(Convert.ToInt32(selectedMode));

                    // Record the ID (GUID) of the new Security Incident
                    var newSecurityIncidentID = svc.Create(newSecurityIncident);

                    newSecurityIncident.Id = newSecurityIncidentID;

                    // Set the file attachment details
                    Guid? fileId = null;

                    string fileName = "My Details.txt";
                    FileInfo fi = new FileInfo(fileName);

                    string fileContents = @"<p>Event details are not provided.</p>";

                    using (FileStream fs = fi.Create())
                    {
                        byte[] data = Encoding.UTF8.GetBytes(fileContents);
                        fs.Write(data, 0, data.Length);
                    }

                    // Upload the file
                    try
                    {
                        fileId = UploadFile(svc,
                           newSecurityIncident.ToEntityReference(),
                           "ts_incidentdetailsattachment",
                           fi,
                           "text/plain");

                        Console.WriteLine($"Uploaded file {fileName}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

                Console.ReadLine();
            }
        }

        // Note: This is from the PowerApps-Samples Repo - https://github.com/microsoft/PowerApps-Samples/blob/9e0e22182ee2e747e29cec4ccdc4399fcf9a8f44/dataverse/orgsvc/C%23-NETCore/FileOperations/Program.cs#L155
        static Guid UploadFile(
                IOrganizationService service,
                EntityReference entityReference,
                string fileAttributeName,
                FileInfo fileInfo,
                string fileMimeType = null,
                int? fileColumnMaxSizeInKb = null)
        {

            // Initialize the upload
            InitializeFileBlocksUploadRequest initializeFileBlocksUploadRequest = new InitializeFileBlocksUploadRequest()
            {
                Target = entityReference,
                FileAttributeName = fileAttributeName,
                FileName = fileInfo.Name
            };

            var initializeFileBlocksUploadResponse =
                (InitializeFileBlocksUploadResponse)service.Execute(initializeFileBlocksUploadRequest);

            string fileContinuationToken = initializeFileBlocksUploadResponse.FileContinuationToken;

            // Capture blockids while uploading
            List<string> blockIds = new List<string>();

            using Stream uploadFileStream = fileInfo.OpenRead();

            int blockSize = 4 * 1024 * 1024; // 4 MB

            byte[] buffer = new byte[blockSize];
            int bytesRead = 0;

            long fileSize = fileInfo.Length;

            if (fileColumnMaxSizeInKb.HasValue && (fileInfo.Length / 1024) > fileColumnMaxSizeInKb.Value)
            {
                throw new Exception($"The file is too large to be uploaded to this column.");
            }


            // The number of iterations that will be required:
            // int blocksCount = (int)Math.Ceiling(fileSize / (float)blockSize);
            int blockNumber = 0;

            // While there is unread data from the file
            while ((bytesRead = uploadFileStream.Read(buffer, 0, buffer.Length)) > 0)
            {
                // The file or final block may be smaller than 4MB
                if (bytesRead < buffer.Length)
                {
                    Array.Resize(ref buffer, bytesRead);
                }

                blockNumber++;

                //string blockId = Convert.ToBase64String(Encoding.UTF8.GetBytes(blockNumber.ToString().PadLeft(16, '0')));
                string blockId = Convert.ToBase64String(Encoding.UTF8.GetBytes(Guid.NewGuid().ToString()));

                blockIds.Add(blockId);

                // Prepare the request
                UploadBlockRequest uploadBlockRequest = new UploadBlockRequest()
                {
                    BlockData = buffer,
                    BlockId = blockId,
                    FileContinuationToken = fileContinuationToken,
                };

                // Send the request
                service.Execute(uploadBlockRequest);
            }

            // Try to get the mimetype if not provided.
            if (string.IsNullOrEmpty(fileMimeType))
            {
                var provider = new FileExtensionContentTypeProvider();

                if (!provider.TryGetContentType(fileInfo.Name, out fileMimeType))
                {
                    fileMimeType = "application/octet-stream";
                }
            }

            // Commit the upload
            CommitFileBlocksUploadRequest commitFileBlocksUploadRequest = new CommitFileBlocksUploadRequest()
            {
                BlockList = blockIds.ToArray(),
                FileContinuationToken = fileContinuationToken,
                FileName = fileInfo.Name,
                MimeType = fileMimeType
            };

            var commitFileBlocksUploadResponse =
                (CommitFileBlocksUploadResponse)service.Execute(commitFileBlocksUploadRequest);

            return commitFileBlocksUploadResponse.FileId;

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

    public static class Choice
    {
        public static List<ChoiceItem> GetChoices(OptionSetMetadata optionSetMetadata)
        {
            List<ChoiceItem> choices = new List<ChoiceItem>();

            foreach (var item in optionSetMetadata.Options)
            {
                choices.Add(new ChoiceItem
                {
                    EnglishLabel = item.Label.LocalizedLabels[0].Label,
                    FrenchLabel = item.Label.LocalizedLabels[1].Label,
                    Value = item.Value
                });
            }

            return choices;
        }
    }

    public class ChoiceItem
    {
        public string EnglishLabel { get; set; }
        public string FrenchLabel { get; set; }
        public int? Value { get; set; }
    }
}
