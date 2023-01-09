using System;
using System.Collections.Generic;
using System.Text;

namespace CSharpAPIDemo_NetCore
{
    public static class LookupColumns
    {
        public struct SecurityIncidentType
        {
            public const string EnglishName = "ts_securityincidenttypenameenglish";
            public const string FrenchName = "ts_securityincidenttypenamefrench";
        }

        public struct TargetElement
        {
            public const string EnglishName = "ts_targetelementnameenglish";
            public const string FrenchName = "ts_targetelementnamefrench";
        }

        // Reporting Company and Stakeholder reside in the same table in ROM
        public struct ReportingCompany
        {
            public const string Name = "name";
        }

        public struct StakeholderOperationType
        {
            public const string EnglishName = "ovs_operationtypenameenglish";
            public const string FrenchName = "ovs_operationtypenamefrench";
        }

        public struct Region
        {
            public const string EnglishName = "ovs_territorynameenglish";
            public const string FrenchName = "ovs_territorynamefrench";
        }

        // Site, SubSite, Origin, and Destination reside in the same table in ROM
        public struct Site
        {
            public const string Name = "msdyn_name";
            public const string Province = "msdyn_stateorprovince";
            public const string Longitude = "msdyn_longitude";
            public const string Latitude = "msdyn_latitude";
        }
    }
}
