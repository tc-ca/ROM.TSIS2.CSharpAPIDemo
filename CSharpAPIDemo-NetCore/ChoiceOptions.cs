using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Text;

namespace CSharpAPIDemo_NetCore
{
    public static class ChoiceOptions
    {
        public struct Program
        {
            public const string SchemaColumnName = "ts_mode";
            public const string GlobalOptionSetName = "ts_securityincidentmode";
        }

        public struct StatusOfRailwayOwner
        {
            public const string SchemaColumnName = "ts_statusofrailwayowner";
            public const string GlobalOptionSetName = "ts_statusofrailwayowner";
        }

        public struct TimeZone
        {
            public const string SchemaColumnName = "ts_timezone";
            public const string GlobalOptionSetName = "ts_timezone";
        }

        public struct Province
        {
            public const string SchemaColumnName = "ts_province";
            public const string GlobalOptionSetName = "ts_province";
        }

        public struct LocationType
        {
            public const string SchemaColumnName = "ts_locationtype";
            public const string GlobalOptionSetName = "ts_locationtype";
        }

        public struct DelaysToOperation
        {
            public const string SchemaColumnName = "ts_delaystooperation";
            public const string GlobalOptionSetName = "ts_delaystooperation";
        }

        public struct PublicOrPrivateCrossing
        {
            public const string SchemaColumnName = "ts_publicorprivatecrossing";
            public const string GlobalOptionSetName = "ts_publicorprivatecrossing";
        }

        public struct RuralOrUrban
        {
            public const string SchemaColumnName = "ts_ruralorurban";
            public const string GlobalOptionSetName = "ts_ruralorurban";
        }

        public struct Injuries
        {
            public const string SchemaColumnName = "ts_injuries";
            public const string GlobalOptionSetName = "ts_injuries";
        }

        public struct Arrests
        {
            public const string SchemaColumnName = "ts_arrests";
            public const string GlobalOptionSetName = "ts_arrestsknownorunknown";
        }

        public struct PoliceResponse
        {
            public const string SchemaColumnName = "ts_policeresponse";
            public const string GlobalOptionSetName = "ts_policeresponse";
        }

        public struct InFlight
        {
            public const string SchemaColumnName = "ts_inflight";
            public const string GlobalOptionSetName = "ts_inflight";
        }
    }
}
