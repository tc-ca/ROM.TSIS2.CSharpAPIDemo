using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROM.TSIS2.CSharpAPIDemo
{
    public static class FetchXMLExamples
    {
        public static string SharePointFileByTableRecordId(string tableRecordId)
        {
            return $@"
                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='ts_sharepointfile'>
                    <attribute name='ts_sharepointfileid' />
                    <attribute name='ts_sharepointfilegroup' />
                    <attribute name='ts_tablerecordid' />
                    <attribute name='ts_tablename' />
                    <attribute name='ts_tablenamefrench' />
                    <link-entity name='ts_sharepointfilegroup' to='ts_sharepointfilegroup' from='ts_sharepointfilegroupid' alias='ts_sharepointfilegroup' link-type='inner' />
                    <filter>
                      <condition attribute='ts_tablerecordid' operator='eq' value='{tableRecordId}' />
                    </filter>
                  </entity>                
                </fetch>            
            ";
        }

        public static string Single_SharePointFile(string tableRecordId)
        {
            return $@"
                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='ts_sharepointfile'>
                    <attribute name='ts_sharepointfilegroup' />
                    <attribute name='ts_sharepointfileid' />
                    <attribute name='ts_tablename' />
                    <attribute name='ts_tablenamefrench' />
                    <attribute name='ts_tablerecordid' />
                    <attribute name='ts_tablerecordname' />
                    <filter>
                      <condition attribute='ts_tablerecordid' operator='eq' value='{tableRecordId}' />
                    </filter>
                  </entity>
                </fetch>            
            ";
        }

        public const string All_Files = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='ts_file'>
                <attribute name='ts_fileid' />
                <attribute name='ts_file' />
                <attribute name='ts_uploadedtosharepoint' />
                <attribute name='ts_exemption' />
                <attribute name='ts_finding' />
                <attribute name='ts_incident' />
                <attribute name='ts_msdyn_workorder' />
                <attribute name='ts_ovs_operation' />
                <attribute name='ts_securityincident' />
                <attribute name='ts_site' />
                <attribute name='ts_stakeholder' />
                <attribute name='ts_workorderservicetask' />
                <attribute name='ts_formintegrationid' />
                <attribute name='ts_description' />
                <attribute name='ts_attachment' />
                <filter type='or'>
                  <condition attribute='ts_uploadedtosharepoint' operator='eq' value='False' />
                  <condition attribute='ts_uploadedtosharepoint' operator='null' />
                </filter>               
                <link-entity name='owner' to='ownerid' from='ownerid' alias='owner' link-type='inner'>
                  <attribute name='name' alias='FileOwner' />
                </link-entity>
                <link-entity name='ts_filecategory' to='ts_filecategory' from='ts_filecategoryid' alias='ts_filecategory' link-type='outer'>
                  <attribute name='ts_categoryen' alias='CategoryEnglish' />
                  <attribute name='ts_categoryfr' alias='CategoryFrench'  />
                </link-entity>
                <link-entity name='ts_filesubcategory' to='ts_filesubcategory' from='ts_filesubcategoryid' alias='ts_filesubcategory' link-type='outer'>
                  <attribute name='ts_filesubcategoryen' alias='SubCategoryEnglish' />
                  <attribute name='ts_filesubcategoryfr' alias='SubCategoryFrench' />
                </link-entity>
              </entity>            
            </fetch>
        ";

        public const string All_Files_Cases = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='ts_files_incidents'>
                <attribute name='incidentid' />
                <attribute name='ts_fileid' />
                <link-entity name='incident' to='incidentid' from='incidentid' alias='incident' link-type='inner'>
                  <attribute name='title' alias='tablerecordname'  />
                </link-entity>
              </entity>            
            </fetch>        
        ";

        public const string All_Files_WorkOrders = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='ts_files_msdyn_workorders'>
                <attribute name='msdyn_workorderid' />
                <attribute name='ts_fileid' />
                <link-entity name='msdyn_workorder' to='msdyn_workorderid' from='msdyn_workorderid' alias='msdyn_workorder' link-type='inner'>
                  <attribute name='msdyn_name' alias='tablerecordname' />
                </link-entity>
              </entity>            
            </fetch>        
        ";

        public const string All_Files_WorkOrderServiceTask = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='ts_files_msdyn_workorderservicetasks'>
                <attribute name='msdyn_workorderservicetaskid' />
                <attribute name='ts_fileid' />
                <link-entity name='msdyn_workorderservicetask' to='msdyn_workorderservicetaskid' from='msdyn_workorderservicetaskid' alias='msdyn_workorderservicetask' link-type='inner'>
                  <attribute name='msdyn_name' alias='tablerecordname' />
                </link-entity>
              </entity>            
            </fetch>        
        ";

        public const string All_Files_Findings = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='ts_ovs_finding_ts_file'>
                <attribute name='ovs_findingid' />
                <attribute name='ts_fileid' />
                <link-entity name='ovs_finding' to='ovs_findingid' from='ovs_findingid' alias='ovs_finding' link-type='inner'>
                  <attribute name='ovs_finding' alias='tablerecordname' />
                </link-entity>
              </entity>            
            </fetch>        
        ";

        public const string All_Files_Sites = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='ts_files_msdyn_functionallocations'>
                <attribute name='msdyn_functionallocationid' />
                <attribute name='ts_fileid' />
                <link-entity name='msdyn_functionallocation' to='msdyn_functionallocationid' from='msdyn_functionallocationid' alias='msdyn_functionallocation' link-type='inner'>
                  <attribute name='ts_functionallocationnameenglish' alias='tablerecordname' />
                </link-entity>
              </entity>            
            </fetch>        
        ";

        public const string All_Files_Operations = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='ts_files_ovs_operations'>
                <attribute name='ovs_operationid' />
                <attribute name='ts_fileid' />
                <link-entity name='ovs_operation' to='ovs_operationid' from='ovs_operationid' alias='ovs_operation' link-type='inner'>
                  <attribute name='ovs_name' alias='tablerecordname' />
                </link-entity>
              </entity>
            </fetch>        
        ";

        public const string All_Files_Stakeholders = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='ts_files_accounts'>
                <attribute name='accountid' />
                <attribute name='ts_fileid' />
                <link-entity name='account' to='accountid' from='accountid' alias='account' link-type='inner'>
                  <attribute name='ovs_legalname' alias='tablerecordname' />
                  <attribute name='name' alias='tablerecordnamebackup' />
                </link-entity>
              </entity>
            </fetch>        
        ";

        public const string All_Security_Incidents = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='ts_securityincident'>
                <attribute name='ts_securityincidentid' />
                <attribute name='ts_name' />
              </entity>
            </fetch>        
        ";

        public const string All_Exemptions = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='ts_exemption'>
                <attribute name='ts_exemptionid' />
                <attribute name='ts_name' />
              </entity>
            </fetch>        
        ";

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

        public const string All_Active_SecurityIncidents = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
              <entity name='ts_securityincident'>
                <attribute name='ts_securityincidentid' />
                <attribute name='ts_name' />
                <filter>
                  <condition attribute='statecode' operator='eq' value='0' />
                </filter>
                <order attribute='ts_name' />
              </entity>
            </fetch>
        ";

        public static string All_Existing_SharePointFiles =
        $@"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='ts_sharepointfile'>
                <attribute name='ts_sharepointfilegroup' />
                <attribute name='ts_sharepointfileid' />
                <attribute name='ts_tablename' />
                <attribute name='ts_tablenamefrench' />
                <attribute name='ts_tablerecordid' />
                <attribute name='ts_tablerecordname' />
                </entity>
            </fetch>            
        ";
        
    }
}
