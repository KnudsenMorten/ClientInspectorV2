# ClientInspector (v2)

## Introduction
**ClientInspector** can be used to collect lots of great information of from your **Windows clients** - and send the data to **Azure LogAnalytics Custom Tables**.

The script collects the following information (settings, information, configuration, state):
1.  User Logged On to Client
2.  Computer information - bios, processor, hardware info, Windows OS info, OS information, last restart
3.  Installed applications, both using WMI and registry
4.  Antivirus Security Center from Windows - default antivirus, state, configuration
5.  Microsoft Defender Antivirus - all settings including ASR, exclusions, realtime protection, etc
6.  Office - version, update channel config, SKUs
7.  VPN client - version, product
8.  LAPS - version
9.  Admin By Request (3rd party) - version
10. Windows Update - last result (when), windows update source information (where), pending updates, last installations (what)
11. Bitlocker - configuration
12. Eventlog - look for specific events including logon events, blue screens, etc.
13. Network adapters - configuration, installed adapters
14. IP information for all adapters
15. Local administrators group membership
16. Windows firewall - settings for all 3 modes
17. Group Policy - last refresh
18. TPM information - relavant to detect machines with/without TPM

Feel free to add more cool data-collections to suit your needs. 
I would love to hear what you are collection. Maybe we can include your ideas into ClientInspector so the whole community can gain access to smart ideas.
Drop me an email on mok@mortenknudsen.net


### How to access the data ?
All the data can be accessed using Kusto (KQL) queries in Azure LogAnalytics - or by the provided Azure Workbooks and Azure Dashboards

<details>
  <summary>Sample query</summary>
  
  ```js
InvClientDefenderAvV2_CL 
| where TimeGenerated > ago(31d)
| summarize CollectionTime = arg_max(CollectionTime, *) by Computer
| where ((AMRunningMode == "Not running") or 
    (parse_version(AMProductVersion) < parse_version("4.18.2203")) or 
    (MPComputerStatusFound == false) or (MPPreferenceFound == false) or
    (RealTimeProtectionEnabled == false) or
    (AntivirusSignatureAge > 7) or (AntispywareSignatureAge > 7) or (NISSignatureAge > 7) or
    (AMRunningMode == "EDR Block Mode") or (AMRunningMode == "Passive Mode") or
    (AntispywareEnabled == false) or
    ((TamperProtectionSource != "ATP") and (TamperProtectionSource != "Intune")) or (IsTamperProtected == false)
    )
| project
    Computer,
    UserLoggedOn,
    CollectionTime,
    MPComputerStatusFound,
    MPPreferenceFound,
    AMEngineVersion,
    AMProductVersion,
    AMRunningMode,
    AMServiceEnabled,
    AMServiceVersion,
    AntispywareEnabled,
    AntispywareSignatureAge,
    AntispywareSignatureLastUpdated,
    AntispywareSignatureVersion,
    AntivirusEnabled,
    AntivirusSignatureAge,
    AntivirusSignatureLastUpdated,
    AntivirusSignatureVersion,
    BehaviorMonitorEnabled,
    DefenderSignaturesOutOfDate,
    DisableAutoExclusions,
    DisableBehaviorMonitoring,
    DisableRealtimeMonitoring,
    DisableScanningMappedNetworkDrivesForFullScan,
    DisableScanningNetworkFiles,
    DisableScriptScanning,
    EnableControlledFolderAccess,
    EnableNetworkProtection,
    FullScanAge,
    IoavProtectionEnabled,
    IsTamperProtected,
    IsVirtualMachine,
    MAPSReporting,
    NISEnabled,
    NISEngineVersion,
    NISSignatureAge,
    NISSignatureLastUpdated,
    NISSignatureVersion,
    OnAccessProtectionEnabled,
    ProductStatus,
    PUAProtection,
    QuickScanAge,
    RealTimeProtectionEnabled,
    RealTimeScanDirection,
    RebootRequired,
    ScanAvgCPULoadFactor,
    SignatureUpdateCatchupInterval,
    SignatureUpdateInterval,
    SubmitSamplesConsent,
    TamperProtectionSource 
  ```
</details>


If you use the [ClientInSpectorV2-DeploymentKit](https://github.com/KnudsenMorten/ClientInspectorV2-DeploymentKit), you will also get access to **13 ready-to-use workbooks** and **14 ready-to-use dashboards**.

If you want to add more views (or workbooks), you can start by investigating the collected data in the custom logs tables using KQL quries. Then make your new views in the workbooks - and pin your favorites to your dashboards.
   
## Archicture & flow
**ClientInspector (v2)** is uploading the collected data into **custom logs** in **Azure LogAnalytics workspace** - using **Log ingestion API**, **Azure Data Collection Rules (DCR)** and **Azure Data Collection Endpoints (DCE)**. 

![Flow](img/flow.png)

The old ClientInspector (v1) was using the HTTP Data Collector API and custom logs (MMA-format).

## How to run ClientInspector ?
You can run the ClientInspector script using your favorite deployment tool. 
Scripts for Microsoft Intune and ConfigMgr (or any other tool running a CMD-file) are provided. 

### How to deploy using Microsoft Intune ?
<details>
  <summary>How to deploy using Microsoft Intune ?</summary>
  
  ```js
  function logSomething(something) {
    console.log('Something', something);
  }
  ```
</details>

### How to deploy using ConfigMgr (or any other tool running a CMD-file) ?



## Infrastructure setup (pre-requisite)
ClientInspector requires some prerequisites to run, which can be deployed using the [ClientInSpectorV2-DeploymentKit](https://github.com/KnudsenMorten/ClientInspectorV2-DeploymentKit)

The following components are needed:

| Azure Resource                | Purpose                                           | More information |
| :-------------                | :-----                                            | :-----           |
| Any REST endpoint             | This is the source sending data                   |                  |
| Log Ingestion API             | This is the new API that replaces the old HTTP Data Collector API | https://learn.microsoft.com/en-us/azure/azure-monitor/logs/logs-ingestion-api-overview |
| Data Collection Endpoint (DCE)| This is the entry point of communication from the endpoint. Data is sent to a **Azure Data Collection Endpoint Ingestion Uri**. A single DCE can support multiple DCRs, so you can specify a different DCR for different sources and target tables. | https://learn.microsoft.com/en-us/azure/azure-monitor/essentials/data-collection-endpoint-overview?tabs=portal |
| Data Collection Rules (DCR)   | Data collection rules define data collected (schema) and specify how and where that data should be sent or stored. The DCR must understand the structure of the input data and the structure of the target table. If the two don't match, it can use a transformation to convert the source data to match the target table. You can also use the transformation to filter source data and perform any other calculations or conversions. | https://learn.microsoft.com/en-us/azure/azure-monitor/essentials/data-collection-rule-overview | 
| Azure LogAnaltyics Workspace  | Data is sent to custom logs in Azure LogAnalytics  | https://learn.microsoft.com/en-us/azure/azure-monitor/logs/log-analytics-overview |
| Azure Workbooks               | As part of the deployment, sample workbooks will be deployed  | https://learn.microsoft.com/en-us/azure/azure-monitor/visualize/workbooks-overview |
| Azure Dashboards              | As part of the deployment, sample workbooks will be deployed  | https://learn.microsoft.com/en-us/azure/azure-monitor/visualize/tutorial-logs-dashboards |
| Kusto (KQL)                   | Data can be analyzed using Kust (KQL) queries | https://learn.microsoft.com/en-us/azure/azure-monitor/logs/get-started-queries |


## Powershell function AzLogDcringestPS (pre-requisite)
ClientInspector requires the Powershell module, **AzLogDcrIngestPS**, also developed by [Morten Knudsen, Microsoft MVP](https://mvp.microsoft.com/en-us/PublicProfile/5005156?fullName=Morten%20Knudsen).

Core features of AzLogDcrIngestPS:
	* create/update the DCRs and tables automatically - based on the source object schema
	* validate the schema for naming convention issues. If exist found, it will mitigate the issues
	* update schema of DCRs and tables, if the structure of the source object changes
	* auto-fix if something goes wrong with a DCR or table
	* can remove data from the source object, if there are colums of data you don't want to send
	* can convert source objects based on CIM or PS objects into PSCustomObjects/array
	* can add relevant information to each record like UserLoggedOn, Computer, CollectionTime

ClientInspector supports to include the Powershell functions in various ways:

You can download latest version here:

[AzLogDcringestPS (Github)](https://github.com/KnudsenMorten/AzLogDcrIngestPS)
  - or -
[AzLogDcringestPS (Powershell Gallery)](https://www.powershellgallery.com/packages/AzLogDcrIngestPS)

### Authentication
Authentication for the Logs Ingestion API is performed at the DCE, which uses standard Azure Resource Manager authentication. 
A common strategy is to use an application ID and application key which is also the method used in ClientInspector.

### Source data
The source data sent by ClientInSpector is formatted in JSON and must match the structure expected by the DCR. 
It doesn't necessarily need to match the structure of the target table because the DCR can include a transformation to convert the data to match the table's structure.

ClientInspector uses several functions within the Powershell module, **AzLogDcIngestPS**, to handle source data adjustsments to **remove "noice" in data**, to **remove prohibited colums in tables/DCR** - and support needs for **transparancy** with extra insight like **UserLoggedOn**, **CollectionTime**, **Computer**:




<details>
  <summary>Sample with usage of functions **Convert-CimArrayToObjectFixStructure**, **Add-CollectionTimeToAllEntriesInArray**, **Add-ColumnDataToAllEntriesInArray**, **ValidateFix-AzLogAnalyticsTableSchemaColumnNames**, **Build-DataArrayToAlignWithSchema**, **Filter-ObjectExcludeProperty**</summary>
  
  ```js
#-------------------------------------------------------------------------------------------
# Collecting data (in)
#-------------------------------------------------------------------------------------------
	
	Write-Output ""
	Write-Output "Collecting Bios information ... Please Wait !"

	$DataVariable = Get-CimInstance -ClassName Win32_BIOS

#-------------------------------------------------------------------------------------------
# Preparing data structure
#-------------------------------------------------------------------------------------------

	# convert CIM array to PSCustomObject and remove CIM class information
	$DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable -Verbose:$Verbose

	# add CollectionTime to existing array
	$DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

	# add Computer & UserLoggedOn info to existing array
	$DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn -Verbose:$Verbose

	# Remove unnecessary columns in schema
	$DataVariable = Filter-ObjectExcludeProperty -Data $DataVariable -ExcludeProperty __*,SystemProperties,Scope,Qualifiers,Properties,ClassPath,Class,Derivation,Dynasty,Genus,Namespace,Path,Property_Count,RelPath,Server,Superclass -Verbose:$Verbose

	# Validating/fixing schema data structure of source data
	$DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

	# Aligning data structure with schema (requirement for DCR)
	$DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
  }
  ```
</details>


You can verify the source object by running this command
````
	# Get insight about the schema structure of an object BEFORE changes. Command is only needed to verify columns in schema
	Get-ObjectSchemaAsArray -Data $DataVariable -Verbose:$Verbose
````


