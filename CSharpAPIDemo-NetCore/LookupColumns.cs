using System;
using System.Collections.Generic;
using System.Text;

namespace CSharpAPIDemo_NetCore
{
    public static class LookupColumns
    {
        //public enum ts_securityincidenttype
        //{
        //    ts_securityincidenttypenameenglish,
        //    ts_securityincidenttypenamefrench
        //}

        public struct SecurityIncidentType
        {
            public const string EnglishName = "ts_securityincidenttypenameenglish";
            public const string FrenchName = "ts_securityincidenttypenamefrench";
        }
    }
}
