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
                              <condition attribute='statecode' operator='eq' value='0' />
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
                              <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            <order attribute='ts_securityincidenttypenamefrench' />
                          </entity>
                        </fetch>
                ";
            }
        }
    }
}
