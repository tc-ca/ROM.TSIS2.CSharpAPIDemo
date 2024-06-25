using Microsoft.AspNetCore.StaticFiles;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;

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
            //Task.WaitAll(Task.Run(async () => await GetTableData()));

            // Connect to the TSIS 2 - ROM API using the CrmServiceClient - Best to use this way
            using (var svc = new CrmServiceClient(connectString))
            {
                // EXAMPLE - Read all retrieved records
                {
                    const string All_Records = @"
                        <fetch xmlns:generator='MarkMpn.SQL4CDS'>
                          <entity name='ts_tcscp'>
                            <attribute name='ts_tcscpid' />
                            <attribute name='ts_name' />
                            <attribute name='ts_processed' />
                            <attribute name='ts_attachment' />
                          </entity>
                        </fetch>                    
                    ";


                    EntityCollection tcscpRecords = svc.RetrieveMultiple(new FetchExpression(All_Records));

                    Console.WriteLine("All records:");
                    Console.WriteLine();

                    foreach (var record in tcscpRecords.Entities)
                    {
                        Console.WriteLine($"ts_tcscpid: {record.Id.ToString()}");
                        Console.WriteLine($"ts_name: {record.GetAttributeValue<String>("ts_name")}");
                        Console.WriteLine($"ts_processed: {record.GetAttributeValue<Boolean>("ts_processed").ToString()}");
                        Console.WriteLine($"ts_attachment: {record.GetAttributeValue<System.Guid>("ts_attachment").ToString()}");
                        Console.WriteLine();
                    }
                }

                // EXAMPLE - Create a record with an attachment
                {
                    Entity newRecord = new Entity("ts_tcscp");

                    // Set the name for the new record
                    newRecord.Attributes["ts_name"] = "My New Record 1";

                    // Record the ID (GUID) of the new record - Note: you have to create a record first before attaching a file
                    var newFileID = svc.Create(newRecord);

                    string fileName = "MyNewFile.json";
                    FileInfo fi = new FileInfo(fileName);

                    // Set the file attachment details
                    Guid? fileId = null;
                    newRecord.Id = newFileID;

                    var myJsonData = new[]
                    {
                        new { name = "Alice", age = 30, city = "Toronto" },
                        new { name = "Bob", age = 25, city = "Vancouver" },
                        new { name = "Charlie", age = 22, city = "Montreal" },
                        new { name = "David", age = 28, city = "Calgary" },
                        new { name = "Eve", age = 35, city = "Edmonton" }
                    };

                    // Convert the list to a JSON string
                    string json = JsonConvert.SerializeObject(myJsonData, Formatting.Indented);

                    string fileContents = json;

                    using (FileStream fs = fi.Create())
                    {
                        byte[] data = Encoding.UTF8.GetBytes(fileContents);
                        fs.Write(data, 0, data.Length);
                    }

                    // Upload the file
                    try
                    {
                        fileId = UploadFile(svc,
                           newRecord.ToEntityReference(),
                           "ts_attachment",
                           fi,
                           "text/plain");

                        Console.WriteLine($"Uploaded file {fileName}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

                // EXAMPLE - Update a record
                {
                    // Get the record you want to update
                    const string UpdateRecordExample = @"
                        <fetch xmlns:generator='MarkMpn.SQL4CDS'>
                          <entity name='ts_tcscp'>
                            <attribute name='createdon' />
                            <attribute name='ts_tcscpid' />
                            <attribute name='ts_name' />
                            <attribute name='ts_processed' />
                            <attribute name='ts_attachment' />
                            <filter>
                              <condition attribute='ts_tcscpid' operator='eq' value='9219075b-2d33-ef11-8e4e-6045bd5d1ea5' />
                            </filter>
                          </entity>
                        </fetch>                    
                    ";

                    EntityCollection record = svc.RetrieveMultiple(new FetchExpression(UpdateRecordExample));
                    var selectedRecord = record[0].Id;

                    var updateRecord = new Entity("ts_tcscp", selectedRecord);

                    // Update a field with some random text
                    updateRecord.Attributes["ts_name"] = "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean";

                    // Also set the processed flag
                    updateRecord.Attributes["ts_processed"] = true;

                    svc.Update(updateRecord);
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

                // Example of a filter being applied while we get all the records
                var response = httpClient.GetAsync(apiPath + "ts_tcscps?$select=ts_tcscpid,ts_name&$filter=ts_name eq 'Test 1'").Result;

                if (response.IsSuccessStatusCode)
                {
                    var jRetrieveResponse = JObject.Parse(response.Content.ReadAsStringAsync().Result);
                    dynamic collAccounts = JsonConvert.DeserializeObject(jRetrieveResponse.ToString());

                    foreach (var data in collAccounts.value)
                    {
                        // Example of writing out the output
                        Console.WriteLine("TCSCP ID: " + data.ts_tcscpid);
                        Console.WriteLine("Site Name: " + data.ts_name);
                        Console.WriteLine("");
                    }
                }
                else
                {
                    return;
                }
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

            Stream uploadFileStream = fileInfo.OpenRead();

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
}
