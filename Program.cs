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
            }
        }
    }
}
