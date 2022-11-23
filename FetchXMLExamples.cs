using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROM.TSIS2.CSharpAPIDemo
{
    public static class FetchXMLExamples
    {
        public const string All_Regions_English = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
              <entity name='territory'>
                <attribute name='territoryid' />
                <attribute name='ovs_territorynameenglish' />
                <order attribute='ovs_territorynameenglish' />
              </entity>
            </fetch>
        ";

        public const string All_Regions_French = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
              <entity name='territory'>
                <attribute name='territoryid' />
                <attribute name='ovs_territorynamefrench' />
                <order attribute='ovs_territorynamefrench' />
              </entity>
            </fetch>
        ";

        public const string All_Operational_AvSec_Stakeholders = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
              <entity name='account'>
                <attribute name='accountid' />
                <attribute name='name' />
                <filter type='or'>
                  <condition attribute='owningbusinessunitname' operator='eq' value='Aviation Security Directorate' />
                  <filter>
                    <condition attribute='owningbusinessunitname' operator='eq' value='Aviation Security Directorate - Domestic' />
                    <condition attribute='ts_stakeholderstatus' operator='eq' value='717750000' />
                  </filter>
                </filter>
                <order attribute='name' />
              </entity>
            </fetch>
        ";

        public const string All_Operational_ISSO_Stakeholders = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
              <entity name='account'>
                <attribute name='accountid' />
                <attribute name='name' />
                <filter>
                  <condition attribute='owningbusinessunitname' operator='eq' value='Intermodal Surface Security Oversight (ISSO)' />
                  <condition attribute='ts_stakeholderstatus' operator='eq' value='717750000' />
                </filter>
                <order attribute='name' />
              </entity>
            </fetch>
        ";
    }
}
