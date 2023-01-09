using System;
using System.Collections.Generic;
using System.Text;

namespace CSharpAPIDemo_NetCore
{
    public static class LookupFetchXML
    {
        public static string SecurityIncidents (bool english = true)
        {
            if (english)
            {
                return @"
                        <fetch>
                          <entity name='ts_securityincidenttype'>
                            <attribute name='ts_securityincidenttypenameenglish' />
                            <filter>
                              <condition attribute='statuscode' operator='eq' value='1' />
                            </filter>
                            <order attribute='ts_securityincidenttypenameenglish' />
                          </entity>
                        </fetch>
                ";
            }
            else
            {
                return @"
                        <fetch>
                          <entity name='ts_securityincidenttype'>
                            <attribute name='ts_securityincidenttypenamefrench' />
                            <filter>
                              <condition attribute='statuscode' operator='eq' value='1' />
                            </filter>
                            <order attribute='ts_securityincidenttypenamefrench' />
                          </entity>
                        </fetch>
                ";
            }
        }

        public static string TargetElements(bool english = true)
        {
            if (english)
            {
                return @"
                        <fetch>
                          <entity name='ts_targetelement'>
                            <attribute name='ts_targetelementnameenglish' />
                            <filter>
                              <condition attribute='statuscode' operator='eq' value='1' />
                            </filter>
                            <order attribute='ts_targetelementnameenglish' />
                          </entity>                        
                        </fetch>
                ";
            }
            else
            {
                return @"
                        <fetch>
                          <entity name='ts_targetelement'>
                            <attribute name='ts_targetelementnamefrench' />
                            <filter>
                              <condition attribute='statuscode' operator='eq' value='1' />
                            </filter>
                            <order attribute='ts_targetelementnamefrench' />
                          </entity>                        
                        </fetch>
                ";
            }
        }

        // Reporting Company and Stakeholder reside in the same table in ROM
        public static string ReportingCompany()
        {
            return @"
                    <fetch>
                        <entity name='account'>
                        <attribute name='name' />
                        <filter>
                            <condition attribute='statuscode' operator='eq' value='1' />
                        </filter>
                        <order attribute='name' />
                        </entity>
                    </fetch>
            ";
        }

        public static string StakeholderOperationType(bool english = true)
        {
            if (english)
            {
                return @"
                        <fetch>
                          <entity name='ovs_operationtype'>
                            <attribute name='ovs_operationtypenameenglish' />
                            <filter>
                              <condition attribute='statuscode' operator='eq' value='1' />
                            </filter>
                            <order attribute='ovs_operationtypenameenglish' />
                          </entity>
                        </fetch>
                ";
            }
            else
            {
                return @"
                        <fetch>
                          <entity name='ovs_operationtype'>
                            <attribute name='ovs_operationtypenamefrench' />
                            <filter>
                              <condition attribute='statuscode' operator='eq' value='1' />
                            </filter>
                            <order attribute='ovs_operationtypenamefrench' />
                          </entity>
                        </fetch>
                ";
            }
        }

        public static string Region(bool english = true)
        {
            if (english)
            {
                return @"
                        <fetch>
                          <entity name='territory'>
                            <attribute name='ovs_territorynameenglish' />
                            <order attribute='ovs_territorynameenglish' />
                          </entity>
                        </fetch>
                ";
            }
            else
            {
                return @"
                        <fetch>
                          <entity name='territory'>
                            <attribute name='ovs_territorynamefrench' />
                            <order attribute='ovs_territorynamefrench' />
                          </entity>
                        </fetch>
                ";
            }
        }

        // Site, SubSite, Origin, and Destination reside in the same table in ROM
        public static string Site()
        {
            return @"
                    <fetch>
                        <entity name='msdyn_functionallocation'>
                        <attribute name='msdyn_name' />
                        <attribute name='msdyn_stateorprovince' />
                        <attribute name='msdyn_longitude' />
                        <attribute name='msdyn_latitude' />
                        <filter>
                            <condition attribute='statuscode' operator='eq' value='1' />
                        </filter>
                        <order attribute='msdyn_name' />
                        </entity>
                    </fetch>
            ";
        }

    }
}
