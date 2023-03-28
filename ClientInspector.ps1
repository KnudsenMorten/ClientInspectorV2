#Requires -RunAsAdministrator
#Requires -Version 5.1

<#
    .NAME
    ClientInspector

    .SYNOPSIS
    This script will collect lots of information from the client - and send the data Azure LogAnalytics Custom Tables.
    The upload happens via Log Ingestion API, Azure Data Collection Rules (DCR) and Azure Data Collection Endpoints.
    
    The script collects the following information (settings, information, configuration, state):
        (1)   User Logged On to Client
        (2)   Computer information - bios, processor, hardware info, Windows OS info, OS information, last restart
        (3)   Installed applications, both using WMI and registry
        (4)   Antivirus Security Center from Windows - default antivirus, state, configuration
        (5)   Microsoft Defender Antivirus - all settings including ASR, exclusions, realtime protection, etc
        (6)   Office - version, update channel config, SKUs
        (7)   VPN client - version, product
        (8)   LAPS - version
        (9)   Admin By Request (3rd party) - version
        (10)  Windows Update - last result (when), windows update source information (where), pending updates, last installations (what)
        (11)  Bitlocker - configuration
        (12)  Eventlog - look for specific events including logon events, blue screens, etc.
        (13)  Network adapters - configuration, installed adapters
        (14)  IP information for all adapters
        (15)  Local administrators group membership
        (16)  Windows firewall - settings for all 3 modes
        (17)  Group Policy - last refresh
        (18)  TPM information - relavant to detect machines with/without TPM
    
    .AUTHOR
    Morten Knudsen, Microsoft MVP - https://mortenknudsen.net

    .LICENSE
    Licensed under the MIT license.

    .PROJECTURI
    https://github.com/KnudsenMorten/ClientInspectorV2

    .EXAMPLE
    .\ClientInspector.ps1 -function:localpath

    .EXAMPLE
    .\ClientInspector.ps1 -function:download

    .EXAMPLE
    .\ClientInspector.ps1 -function:localpath -verbose:$true

    .EXAMPLE
    .\ClientInspector.ps1 -verbose:$false -function:psgallery -Scope:currentuser

    .WARRANTY
    Use at your own risk, no warranty given!
#>

param(
      [parameter(Mandatory=$false)]
          [ValidateSet("Download","LocalPath","DevMode","PsGallery")]
          [string]$Function = "PsGallery",        # it will default to download if not specified
      [parameter(Mandatory=$false)]
          [ValidateSet("CurrentUser","AllUsers")]
          [string]$Scope = "CurrentUser"        # it will default to download if not specified
     )

Write-Output ""
Write-Output "ClientInspector | Inventory of Operational & Security-related information"
Write-Output "Developed by Morten Knudsen, Microsoft MVP"
Write-Output ""
  
##########################################
# VARIABLES
##########################################

<# ----- onboarding lines ----- BEGIN #>






# On some computers doing a WMI query after applications is REALLY slow. When testing, it is nice to be able to disable this collection temporarely
$Collect_Applications_WMI                   = $true

# This variable can be used to set verbose mode-flag (true/false), if you test using Powershell ISE. 
# Normally you will set the verbose-flag on the commandline
# Remember to remove the # in front of $Verbose to activate it. 
# Remember to add the # in front of #verbose, when you want to control it through the commandline

# $Verbose                                    = $true

<# ----- onboarding lines ----- END  #>

$LastRun_RegPath                            = "HKLM:\SOFTWARE\ClientInspector"
$LastRun_RegKey                             = "ClientInspector_System"

# default variables
$DNSName                                    = (Get-CimInstance win32_computersystem).DNSHostName +"." + (Get-CimInstance win32_computersystem).Domain
$ComputerName                               = (Get-CimInstance win32_computersystem).DNSHostName
[datetime]$CollectionTime                   = ( Get-date ([datetime]::Now.ToUniversalTime()) -format "yyyy-MM-ddTHH:mm:ssK" )

# script run mode - normal or verbose
If ( ($psBoundParameters['verbose'] -eq $true) -or ($verbose -eq $true) )
    {
        Write-Output "Verbose mode ON"
        $global:Verbose = $true
        $VerbosePreference = "Continue"  # Stop, Inquire, Continue, SilentlyContinue
    }
Else
    {
        $global:Verbose = $false
        $VerbosePreference = "SilentlyContinue"  # Stop, Inquire, Continue, SilentlyContinue
    }


############################################################################################################################################
# FUNCTIONS
############################################################################################################################################

    $PowershellVersion  = [version]$PSVersionTable.PSVersion
    If ([Version]$PowershellVersion -ge "5.1")
        {
            $PS_WMF_Compliant  = $true
            $EnableUploadViaLogHub  = $false
        }
    Else
        {
            $PS_WMF_Compliant  = $false
            $EnableUploadViaLogHub  = $true
            Import-module "$($LogHubPsModulePath)\AzLogDcrIngestPS.psm1" -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
        }

    # directory where the script was started
    $ScriptDirectory = $PSScriptRoot

    switch ($Function)
        {   
            "Download"
                {
                    # force download using Github. This is needed for Intune remediations, since the functions library are large, and Intune only support 200 Kb at the moment
                    Write-Output "Downloading latest version of module AzLogDcrIngestPS from https://github.com/KnudsenMorten/AzLogDcrIngestPS"
                    Write-Output "into local path $($ScriptDirectory)"

                    # delete existing file if found to download newest version
                    If (Test-Path "$($ScriptDirectory)\AzLogDcrIngestPS.psm1")
                        {
                            Remove-Item -Path "$($ScriptDirectory)\AzLogDcrIngestPS.psm1"
                        }

                     # download newest version
                    $Download = (New-Object System.Net.WebClient).DownloadFile("https://raw.githubusercontent.com/KnudsenMorten/AzLogDcrIngestPS/main/AzLogDcrIngestPS.psm1", "$($ScriptDirectory)\AzLogDcrIngestPS.psm1")
                    
                    Start-Sleep -s 3
                    
                    # load file if found - otherwise terminate
                    If (Test-Path "$($ScriptDirectory)\AzLogDcrIngestPS.psm1")
                        {
                            Import-module "$($ScriptDirectory)\AzLogDcrIngestPS.psm1" -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
                        }
                    Else
                        {
                            Write-Output "Powershell module AzLogDcrIngestPS was NOT found .... terminating !"
                            break
                        }
                }

            "PsGallery"
                {
                        # check for AzLogDcrIngestPS
                            $ModuleCheck = Get-Module -Name AzLogDcrIngestPS -ListAvailable -ErrorAction SilentlyContinue
                            If (!($ModuleCheck))
                                {
                                    # check for NuGet package provider
                                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

                                    Write-Output ""
                                    Write-Output "Checking Powershell PackageProvider NuGet ... Please Wait !"
                                        if (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue -WarningAction SilentlyContinue) 
                                            {
                                                Write-Host "OK - PackageProvider NuGet is installed"
                                            } 
                                        else 
                                            {
                                                try
                                                    {
                                                        Write-Host "Installing NuGet package provider .. Please Wait !"
                                                        Install-PackageProvider -Name NuGet -Scope $Scope -Confirm:$false -Force
                                                    }
                                                catch [Exception] {
                                                    $_.message 
                                                    exit
                                                }
                                            }

                                    Write-Output "Powershell module AzLogDcrIngestPS was not found !"
                                    Write-Output "Installing latest version from PsGallery in scope $Scope .... Please Wait !"

                                    Install-module -Name AzLogDcrIngestPS -Repository PSGallery -Force -Scope $Scope
                                    import-module -Name AzLogDcrIngestPS -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
                                }

                            Elseif ($ModuleCheck)
                                {
                                    # sort to get highest version, if more versions are installed
                                    $ModuleCheck = Sort-Object -Descending -Property Version -InputObject $ModuleCheck
                                    $ModuleCheck = $ModuleCheck[0]

                                    Write-Output "Checking latest version at PsGallery for AzLogDcrIngestPS module"
                                    $online = Find-Module -Name AzLogDcrIngestPS -Repository PSGallery

                                    #compare versions
                                    if ( ([version]$online.version) -gt ([version]$ModuleCheck.version) ) 
                                        {
                                            Write-Output "Newer version ($($online.version)) detected"
                                            Write-Output "Updating AzLogDcrIngestPS module .... Please Wait !"
                                            Update-module -Name AzLogDcrIngestPS -Force
                                            import-module -Name AzLogDcrIngestPS -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
                                        }
                                    else
                                        {
                                            # No new version detected ... continuing !
                                            Write-Output "OK - Running latest version"
                                            $UpdateAvailable = $False
                                            import-module -Name AzLogDcrIngestPS -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
                                        }
                                }
                }
            "LocalPath"        # Typucaly used in ConfigMgr environment (or similar) where you run the script locally
                {
                    If (Test-Path "$($ScriptDirectory)\AzLogDcrIngestPS.psm1")
                        {
                            Write-Output "Using AzLogDcrIngestPS module from local path $($ScriptDirectory)"
                            Import-module "$($ScriptDirectory)\AzLogDcrIngestPS.psm1" -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
                        }
                    Else
                        {
                            Write-Output "Required Powershell function was NOT found .... terminating !"
                            Exit
                        }
                }
        }


###############################################################
# Global Variables
#
# Used to mitigate throttling in Azure Resource Graph
# Needs to be loaded after load of functions
###############################################################

    # building global variable with all DCEs, which can be viewed by Log Ingestion app
    $global:AzDceDetails = Get-AzDceListAll -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose
    
    # building global variable with all DCRs, which can be viewed by Log Ingestion app
    $global:AzDcrDetails = Get-AzDcrListAll -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


############################################################################################################################################
# MAIN PROGRAM
############################################################################################################################################

    #-------------------------------------------------------------------------------------------------------------
    # Initial Powershell module check
    #-------------------------------------------------------------------------------------------------------------

        $ModuleCheck = Get-Module -Name PSWindowsUpdate -ListAvailable -ErrorAction SilentlyContinue
        If (!($ModuleCheck))
            {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

                Write-Output ""
                Write-Output "Checking Powershell PackageProvider NuGet ... Please Wait !"
                    if (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue -WarningAction SilentlyContinue) 
                        {
                            Write-Host "  OK - PackageProvider NuGet is installed"
                        } 
                    else 
                        {
                            try {
                                Install-PackageProvider -Name NuGet -Scope AllUsers -Confirm:$false -Force
                            }
                            catch [Exception] {
                                $_.message 
                                exit
                            }
                        }

                Write-Output ""
                Write-Output "Checking Powershell Module PSWindowsUpdate ... Please Wait !"
                    if (Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue -WarningAction SilentlyContinue) 
                        {
                            Write-output "  OK - Powershell Modue PSWindowsUpdate is installed"
                        } 
                    else 
                        {
                            try {
                                Write-Output "  Installing Powershell Module PSWindowsUpdate .... Please Wait !"
                                Install-Module -Name PSWindowsUpdate -AllowClobber -Scope AllUsers -Confirm:$False -Force
                                Import-Module -Name PSWindowsUpdate
                            }
                            catch [Exception] {
                                $_.message 
                                exit
                            }
                        }
            }


###############################################################
# USER [1]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "User information [1]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------

        $TableName = 'InvClientComputerUserLoggedOnV2'
        $DcrName   = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------

        Write-Output ""
        Write-Output "Collecting User information ... Please Wait !"

        $UserLoggedOnRaw = Get-Process -IncludeUserName -Name explorer | Select-Object UserName -Unique
        $UserLoggedOn    = $UserLoggedOnRaw.UserName

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        # Build array
        $DataVariable = [pscustomobject]@{
                                            UserLoggedOn         = $UserLoggedOn
                                         }

        # add CollectionTime to existing array
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

        # add Computer & ComputerFqdn & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Verbose:$Verbose

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose

    #-------------------------------------------------------------------------------------------
    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
    #-------------------------------------------------------------------------------------------

        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                            -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                            -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                            -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                            -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                            -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                            -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine

    #-----------------------------------------------------------------------------------------------
    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
    #-----------------------------------------------------------------------------------------------

        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


###############################################################
# COMPUTER INFORMATION [2]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "COMPUTER INFORMATION [2]"
    Write-output ""

    ####################################
    # Bios
    ####################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientComputerInfoBiosV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting Bios information ... Please Wait !"

            $DataVariable = Get-CimInstance -ClassName Win32_BIOS

            If ($DataVariable)
                {
                    #-------------------------------------------------------------------------------------------
                    # Preparing data structure
                    #-------------------------------------------------------------------------------------------
                        # convert CIM array to PSCustomObject and remove CIM class information
                        $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable -Verbose:$Verbose
    
                        # add CollectionTime to existing array
                        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                        # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                        # Validating/fixing schema data structure of source data
                        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                        # Aligning data structure with schema (requirement for DCR)
                        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose

                    #-------------------------------------------------------------------------------------------
                    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
                    #-------------------------------------------------------------------------------------------

                        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
                    #-----------------------------------------------------------------------------------------------
                    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
                    #-----------------------------------------------------------------------------------------------

                        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose
                } # If $DataVariable


    ####################################
    # Processor
    ####################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientComputerInfoProcessorV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting Processor information ... Please Wait !"
            $DataVariable = Get-CimInstance -ClassName Win32_Processor | Select-Object -ExcludeProperty "CIM*"

            If ($DataVariable)
                {
                    #-------------------------------------------------------------------------------------------
                    # Preparing data structure
                    #-------------------------------------------------------------------------------------------

                        # convert CIM array to PSCustomObject and remove CIM class information
                        $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable -Verbose:$Verbose
    
                        # add CollectionTime to existing array
                        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                        # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                        # Validating/fixing schema data structure of source data
                        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                        # Aligning data structure with schema (requirement for DCR)
                        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose

                    #-------------------------------------------------------------------------------------------
                    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
                    #-------------------------------------------------------------------------------------------

                        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
                    #-----------------------------------------------------------------------------------------------
                    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
                    #-----------------------------------------------------------------------------------------------

                        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose

                } # If $DataVariable

    ####################################
    # Computer System
    ####################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientComputerInfoSystemV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting Computer system information ... Please Wait !"

            $DataVariable = Get-CimInstance -ClassName Win32_ComputerSystem

            If ($DataVariable)
                {

                    #-------------------------------------------------------------------------------------------
                    # Preparing data structure
                    #-------------------------------------------------------------------------------------------

                        # convert CIM array to PSCustomObject and remove CIM class information
                        $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable -Verbose:$Verbose
    
                        # add CollectionTime to existing array
                        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                        # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                        # Validating/fixing schema data structure of source data
                        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                        # Aligning data structure with schema (requirement for DCR)
                        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose

                    #-------------------------------------------------------------------------------------------
                    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
                    #-------------------------------------------------------------------------------------------

                        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
                    #-----------------------------------------------------------------------------------------------
                    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
                    #-----------------------------------------------------------------------------------------------

                        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose

                } # If $DataVariable


    ####################################
    # Computer Info
    ####################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientComputerInfoV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting computer information ... Please Wait !"

            $DataVariable = Get-ComputerInfo

            If ($DataVariable)
                {
                    #-------------------------------------------------------------------------------------------
                    # Preparing data structure
                    #-------------------------------------------------------------------------------------------

                        # convert CIM array to PSCustomObject and remove CIM class information
                        $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable -Verbose:$Verbose
    
                        # add CollectionTime to existing array
                        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                        # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                        # Validating/fixing schema data structure of source data
                        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                        # Aligning data structure with schema (requirement for DCR)
                        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose

                    #-------------------------------------------------------------------------------------------
                    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
                    #-------------------------------------------------------------------------------------------

                        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
                    #-----------------------------------------------------------------------------------------------
                    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
                    #-----------------------------------------------------------------------------------------------

                        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose
                } # If $DataVariable


    ####################################
    # OS Info
    ####################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientComputerOSInfoV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting OS information ... Please Wait !"

            $DataVariable = Get-CimInstance -ClassName Win32_OperatingSystem

            If ($DataVariable)
                {
                    #-------------------------------------------------------------------------------------------
                    # Preparing data structure
                    #-------------------------------------------------------------------------------------------

                        # convert CIM array to PSCustomObject and remove CIM class information
                        $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable -Verbose:$Verbose
    
                        # add CollectionTime to existing array
                        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                        # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                        # Validating/fixing schema data structure of source data
                        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                        # Aligning data structure with schema (requirement for DCR)
                        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose

                    #-------------------------------------------------------------------------------------------
                    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
                    #-------------------------------------------------------------------------------------------

                        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
                    #-----------------------------------------------------------------------------------------------
                    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
                    #-----------------------------------------------------------------------------------------------

                        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose

                } # If $DataVariable


    ####################################
    # Last Restart
    ####################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientComputerInfoLastRestartV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting Last restart information ... Please Wait !"

            $LastRestart = Get-CimInstance -ClassName win32_operatingsystem | Select lastbootuptime
            $LastRestart = (Get-date $LastRestart.LastBootUpTime)

            $Today = (GET-DATE)
            $TimeSinceLastReboot = NEW-TIMESPAN –Start $LastRestart –End $Today

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------

            $DataVariable = [pscustomobject]@{
                                                LastRestart          = $LastRestart
                                                DaysSinceLastRestart = $TimeSinceLastReboot.Days
                                                }

            # add CollectionTime to existing array
            $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

            # add Computer, ComputerFqdn & UserLoggedOn info to existing array
            $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

            # Validating/fixing schema data structure of source data
            $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

            # Aligning data structure with schema (requirement for DCR)
            $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose


        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                               -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                               -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                               -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                               -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                               -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                             -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose



###############################################################
# APPLICATIONS (WMI) [3]
###############################################################

    If ($Collect_Applications_WMI -eq $true)
        {
            Write-output ""
            Write-output "#########################################################################################"
            write-output "INSTALLED APPLICATIONS INFORMATION [3]"
            Write-output ""

            #------------------------------------------------
            # Installed Application (WMI)
            #------------------------------------------------

                #-------------------------------------------------------------------------------------------
                # Variables
                #-------------------------------------------------------------------------------------------
            
                    $TableName  = 'InvClientApplicationsFromWmiV2'   # must not contain _CL
                    $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

                #-------------------------------------------------------------------------------------------
                # Collecting data (in)
                #-------------------------------------------------------------------------------------------
            
                    Write-Output "Collecting installed application information via WMI (slow) ... Please Wait !"

                    $DataVariable = Get-CimInstance -Class Win32_Product

                    If ($DataVariable)
                        {

                            #-------------------------------------------------------------------------------------------
                            # Preparing data structure
                            #-------------------------------------------------------------------------------------------
    
                                # convert Cim object and remove PS class information
                                $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable -Verbose:$Verbose

                                # add CollectionTime to existing array
                                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                                # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                                # Get insight about the schema structure of an object BEFORE changes. Command is only needed to verify columns in schema
                                # $SchemaBefore = Get-ObjectSchemaAsArray -Data $DataVariable -Verbose:$Verbose
        
                                # Remove unnecessary columns in schema
                                $DataVariable = Filter-ObjectExcludeProperty -Data $DataVariable -ExcludeProperty __*,SystemProperties,Scope,Qualifiers,Properties,ClassPath,Class,Derivation,Dynasty,Genus,Namespace,Path,Property_Count,RelPath,Server,Superclass -Verbose:$Verbose

                                # Validating/fixing schema data structure of source data
                                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                                # Aligning data structure with schema (requirement for DCR)
                                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose


                            #-------------------------------------------------------------------------------------------
                            # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
                            #-------------------------------------------------------------------------------------------

                                $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                                                   -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                                                   -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                                                   -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                                                   -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                                                   -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                                                   -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
                            #-----------------------------------------------------------------------------------------------
                            # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
                            #-----------------------------------------------------------------------------------------------

                                $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName -BatchAmount 1 `
                                                                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose
                    
                    } # If $DataVariable


        }


###############################################################
# APPLICATIONS (REGISTRY) [3]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    write-output "INSTALLED APPLICATIONS INFORMATION [3]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientApplicationsFromRegistryV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting installed applications information via registry ... Please Wait !"

        $UninstallValuesX86 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue
        $UninstallValuesX64 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue

        $DataVariable       = $UninstallValuesX86
        $DataVariable      += $UninstallValuesX64

        If ($DataVariable)
            {

                #-------------------------------------------------------------------------------------------
                # Preparing data structure
                #-------------------------------------------------------------------------------------------


                    # removing apps without DisplayName fx KBs
                    Try
                        {
                            $DataVariable = $DataVariable | Where-Object { $_.DisplayName -ne $null } -ErrorAction SilentlyContinue
                        }
                    Catch
                        {
                        }

                    # convert PS object and remove PS class information
                    $DataVariable = Convert-PSArrayToObjectFixStructure -Data $DataVariable -Verbose:$Verbose

                    # add CollectionTime to existing array
                    $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                    # add Computer & ComputerFqdn info to existing array
                    $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Verbose:$Verbose

                    # Get insight about the schema structure of an object BEFORE changes. Command is only needed to verify columns in schema
                    # $SchemaBefore = Get-ObjectSchemaAsArray -Data $DataVariable
        
                    # Remove unnecessary columns in schema
                    $DataVariable = Filter-ObjectExcludeProperty -Data $DataVariable -ExcludeProperty Memento*,Inno*,'(default)',1033 -Verbose:$Verbose

                    # Validating/fixing schema data structure of source data
                    $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                    # Aligning data structure with schema (requirement for DCR)
                    $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose

                    # set variable to be used later in script ($InstalledApplications)
                    $InstalledApplications = $DataVariable


                #-------------------------------------------------------------------------------------------
                # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
                #-------------------------------------------------------------------------------------------

                    $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                                       -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                                       -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                                       -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                                       -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                                       -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                                       -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
                #-----------------------------------------------------------------------------------------------
                # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
                #-----------------------------------------------------------------------------------------------

                    $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName -BatchAmount 1 `
                                                                                     -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose

            
            } # If $DataVariable


###############################################################
# ANTIVIRUS SECURITY CENTER [4]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "ANTIVIRUS INFORMATION SECURITY CENTER [4]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientAntivirusV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting antivirus information ... Please Wait !"

        $PrimaryAntivirus                           = 'NOT FOUND'
        $Alternative1Antivirus                      = 'NOT FOUND'
        $Alternative2Antivirus                      = 'NOT FOUND'
        $PrimaryAntivirusProduct                    = ""
        $PrimaryAntivirusEXE                        = ""
        $PrimaryAntivirusProductStateCode           = ""
        $PrimaryAntivirusProductStateTimestamp      = ""
        $PrimaryAntivirusDefinitionStatus           = ""
        $PrimaryAntivirusRealTimeStatus             = ""
        $Alternative1AntivirusProduct               = ""
        $Alternative1AntivirusEXE                   = ""
        $Alternative1AntivirusProductStateCode      = ""
        $Alternative1AntivirusProductStateTimestamp = ""
        $Alternative1AntivirusDefinitionStatus      = ""
        $Alternative1AntivirusRealTimeStatus        = ""
        $Alternative2AntivirusProduct               = ""
        $Alternative2AntivirusEXE                   = ""
        $Alternative2AntivirusProductStateCode      = ""
        $Alternative2AntivirusProductStateTimestamp = ""
        $Alternative2AntivirusDefinitionStatus      = ""
        $Alternative2AntivirusRealTimeStatus        = ""

        $AntiVirusProducts = Get-CimInstance -Namespace "root\SecurityCenter2" -Class AntiVirusProduct

        $ret = @()
        foreach($AntiVirusProduct in $AntiVirusProducts)
            {
                switch ($AntiVirusProduct.productState) 
                    {
                        "262144" {$defstatus = "Up to date" ;$rtstatus = "Disabled"}
                        "262160" {$defstatus = "Out of date" ;$rtstatus = "Disabled"}
                        "266240" {$defstatus = "Up to date" ;$rtstatus = "Enabled"}
                        "266256" {$defstatus = "Out of date" ;$rtstatus = "Enabled"}
                        "393216" {$defstatus = "Up to date" ;$rtstatus = "Disabled"}
                        "393232" {$defstatus = "Out of date" ;$rtstatus = "Disabled"}
                        "393488" {$defstatus = "Out of date" ;$rtstatus = "Disabled"}
                        "397312" {$defstatus = "Up to date" ;$rtstatus = "Enabled"}
                        "397328" {$defstatus = "Out of date" ;$rtstatus = "Enabled"}
                        "397584" {$defstatus = "Out of date" ;$rtstatus = "Enabled"}
                        "397568" {$defstatus = "Up to date" ;$rtstatus = "Enabled"}          # Windows Defender
                        "393472" {$defstatus = "Up to date" ;$rtstatus = "Disabled"}         # Windows Defender
                        "397584" {$defstatus = "Out of date" ;$rtstatus = "Enabled"}         # Windows Defender
                        default {$defstatus = "Unknown" ;$rtstatus = "Unknown"}
                    }

                    # Detect Primary
                    If ( ($rtstatus -eq 'Enabled') -and ($PrimaryAntivirus -eq 'NOT FOUND') )
                        {
                            $PrimaryAntivirusProduct = $AntiVirusProduct.displayName
                            $PrimaryAntivirusEXE = $AntiVirusProduct.pathToSignedReportingExe
                            $PrimaryAntivirusProductStateCode = $AntiVirusProduct.productState
                            $PrimaryAntivirusProductStateTimestamp = $AntiVirusProduct.timestamp
                            $PrimaryAntivirusDefinitionStatus = $DefStatus
                            $PrimaryAntivirusRealTimeStatus = $rtstatus
                            $PrimaryAntivirus = 'FOUND'
                        }
        
                    # Detect Alternative 1
                    If ( ($rtstatus -eq 'disabled') -and ($Alternative1Antivirus -eq 'NOT FOUND') )
                        {
                            $Alternative1AntivirusProduct = $AntiVirusProduct.displayName
                            $Alternative1AntivirusEXE = $AntiVirusProduct.pathToSignedReportingExe
                            $Alternative1AntivirusProductStateCode = $AntiVirusProduct.productState
                            $Alternative1AntivirusProductStateTimestamp = $AntiVirusProduct.timestamp
                            $Alternative1AntivirusDefinitionStatus = $DefStatus
                            $Alternative1AntivirusRealTimeStatus = $rtstatus
                            $Alternative1Antivirus = 'FOUND'
                        }

                    # Detect Alternative 2
                    If ( ($rtstatus -eq 'disabled') -and ($Alternative1Antivirus -eq 'FOUND') -eq ($Alternative2Antivirus -eq 'NOT FOUND') )
                        {
                            $Alternative1AntivirusProduct = $AntiVirusProduct.displayName
                            $Alternative1AntivirusEXE = $AntiVirusProduct.pathToSignedReportingExe
                            $Alternative1AntivirusProductStateCode = $AntiVirusProduct.productState
                            $Alternative1AntivirusProductStateTimestamp = $AntiVirusProduct.timestamp
                            $Alternative1AntivirusDefinitionStatus = $DefStatus
                            $Alternative1AntivirusRealTimeStatus = $rtstatus
                            $Alternative1Antivirus = 'FOUND'
                        }
            }


    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        $DataVariable  = [pscustomobject]@{
                                            PrimaryAntivirusProduct = $PrimaryAntivirusProduct
                                            PrimaryAntivirusEXE = $PrimaryAntivirusEXE
                                            PrimaryAntivirusProductStateCode = $PrimaryAntivirusProductStateCode
                                            PrimaryAntivirusProductStateTimestamp = $PrimaryAntivirusProductStateTimestamp
                                            PrimaryAntivirusDefinitionStatus = $PrimaryAntivirusDefinitionStatus
                                            PrimaryAntivirusRealTimeStatus = $PrimaryAntivirusRealTimeStatus
                                            Alternative1AntivirusProduct = $Alternative1AntivirusProduct
                                            Alternative1AntivirusEXE = $Alternative1AntivirusEXE
                                            Alternative1AntivirusProductStateCode = $Alternative1AntivirusProduct
                                            Alternative1AntivirusProductStateTimestamp = $Alternative1AntivirusProductStateTimestamp
                                            Alternative1AntivirusDefinitionStatus = $Alternative1AntivirusDefinitionStatus
                                            Alternative1AntivirusRealTimeStatus = $Alternative1AntivirusRealTimeStatus
                                            Alternative2AntivirusProduct = $Alternative2AntivirusProduct
                                            Alternative2AntivirusEXE = $Alternative2AntivirusEXE
                                            Alternative2AntivirusProductStateCode = $Alternative2AntivirusProduct
                                            Alternative2AntivirusProductStateTimestamp = $Alternative2AntivirusProductStateTimestamp
                                            Alternative2AntivirusDefinitionStatus = $Alternative2AntivirusDefinitionStatus
                                            Alternative2AntivirusRealTimeStatus = $Alternative2AntivirusRealTimeStatus
                                          }
    
        # add CollectionTime to existing array
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

        # add Computer, ComputerFqdn & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose

    #-------------------------------------------------------------------------------------------
    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
    #-------------------------------------------------------------------------------------------

        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
    #-----------------------------------------------------------------------------------------------
    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
    #-----------------------------------------------------------------------------------------------

        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


###############################################################
# MICROSOFT DEFENDER ANTIVIRUS [5]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "MICROSOFT DEFENDER ANTIVIRUS INFORMATION [5]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientDefenderAvV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting Microsoft Defender Antivirus information ... Please Wait !"

        Try
            {
                $MPComputerStatus = Get-MpComputerStatus
            }
        Catch
            {
                $MPComputerStatus = $null
            }


        Try
            {
                $MPPreference = Get-MpPreference
            }
        Catch
            {
                $MPPreference = $null
            }


    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        # Defender was found !
        If ($MPComputerStatus) 
            {
                $DefenderObject = new-object PSCustomObject
                
                # MPComputerStatus
                    $DefenderObject | add-member -MemberType NoteProperty -Name MPComputerStatusFound -Value $True

                    $ObjColumns = ($MPComputerStatus | get-member -MemberType Property)
                    ForEach ($Entry in $MPComputerStatus)
                        {
                            ForEach ($Column in $ObjColumns)
                                {
                                    $ColumnName = $Column.name
                                    $DefenderObject | add-member -MemberType NoteProperty -Name $ColumnName -Value $Entry.$ColumnName -force
                                }
                        }

                # MPPreference
                    $DefenderObject | add-member -MemberType NoteProperty -Name MPPreferenceFound -Value $True

                    $ObjColumns = ($MPPreference | get-member -MemberType Property)
                    ForEach ($Entry in $MPPreference)
                        {
                            ForEach ($Column in $ObjColumns)
                                {
                                    $ColumnName = $Column.name
                                    $DefenderObject | add-member -MemberType NoteProperty -Name $ColumnName -Value $Entry.$ColumnName -Force
                                }
                        }
            }
        Else    # no defender was found !
            {
                # Empty Defender info - modules not found !!
                $DefenderObject = new-object PSCustomObject

                # MPComputerStatus
                    $DefenderObject | add-member -MemberType NoteProperty -Name MPComputerStatusFound -Value $False

                    $MPComputerStatusArray = @( "AMEngineVersion", `
                                                "AMProductVersion", `
                                                "AMRunningMode", `
                                                "AMServiceEnabled", `
                                                "AMServiceVersion", `
                                                "AntispywareEnabled", `
                                                "AntispywareSignatureAge", `
                                                "AntispywareSignatureLastUpdated", `
                                                "AntispywareSignatureVersion", `
                                                "AntivirusEnabled", `
                                                "AntivirusSignatureAge", `
                                                "AntivirusSignatureLastUpdated", `
                                                "AntivirusSignatureVersion", `
                                                "BehaviorMonitorEnabled", `
                                                "ComputerID", `
                                                "ComputerState", `
                                                "DefenderSignaturesOutOfDate", `
                                                "DeviceControlDefaultEnforcement", `
                                                "DeviceControlPoliciesLastUpdated", `
                                                "DeviceControlState", `
                                                "FullScanAge", `
                                                "FullScanEndTime", `
                                                "FullScanOverdue", `
                                                "FullScanRequired", `
                                                "FullScanSignatureVersion", `
                                                "FullScanStartTime", `
                                                "IoavProtectionEnabled", `
                                                "IsTamperProtected", `
                                                "IsVirtualMachine", `
                                                "LastFullScanSource", `
                                                "LastQuickScanSource", `
                                                "NISEnabled", `
                                                "NISEngineVersion", `
                                                "NISSignatureAge", `
                                                "NISSignatureLastUpdated", `
                                                "NISSignatureVersion", `
                                                "OnAccessProtectionEnabled", `
                                                "ProductStatus", `
                                                "PSComputerName", `
                                                "QuickScanAge", `
                                                "QuickScanEndTime", `
                                                "QuickScanOverdue", `
                                                "QuickScanSignatureVersion", `
                                                "QuickScanStartTime", `
                                                "RealTimeProtectionEnabled", `
                                                "RealTimeScanDirection", `
                                                "RebootRequired", `
                                                "SmartAppControlExpiration", `
                                                "SmartAppControlState", `
                                                "TamperProtectionSource", `
                                                "TDTMode", `
                                                "TDTSiloType", `
                                                "TDTStatus", `
                                                "TDTTelemetry", `
                                                "TroubleShootingDailyMaxQuota", `
                                                "TroubleShootingDailyQuotaLeft", `
                                                "TroubleShootingEndTime", `
                                                "TroubleShootingExpirationLeft", `
                                                "TroubleShootingMode", `
                                                "TroubleShootingModeSource", `
                                                "TroubleShootingQuotaResetTime", `
                                                "TroubleShootingStartTime"
                                               )

                # loop
                ForEach ($Column in $MPComputerStatusArray)
                    {
                        $DefenderObject | add-member -MemberType NoteProperty -Name $Column -Value "" -Force
                    }


                # MPPreference
                    $DefenderObject | add-member -MemberType NoteProperty -Name MPPreferenceFound -Value $False

                    $MPPreferenceArray = @( "AllowDatagramProcessingOnWinServer", `
                                            "AllowNetworkProtectionDownLevel", `
                                            "AllowNetworkProtectionOnWinServer", `
                                            "AllowSwitchToAsyncInspection", `
                                            "AttackSurfaceReductionOnlyExclusions", `
                                            "AttackSurfaceReductionRules_Actions", `
                                            "AttackSurfaceReductionRules_Ids", `
                                            "CheckForSignaturesBeforeRunningScan", `
                                            "CloudBlockLevel", `
                                            "CloudExtendedTimeout", `
                                            "ComputerID", `
                                            "ControlledFolderAccessAllowedApplications", `
                                            "ControlledFolderAccessProtectedFolders", `
                                            "DefinitionUpdatesChannel", `
                                            "DisableArchiveScanning", `
                                            "DisableAutoExclusions", `
                                            "DisableBehaviorMonitoring", `
                                            "DisableBlockAtFirstSeen", `
                                            "DisableCatchupFullScan", `
                                            "DisableCatchupQuickScan", `
                                            "DisableCpuThrottleOnIdleScans", `
                                            "DisableDatagramProcessing", `
                                            "DisableDnsOverTcpParsing", `
                                            "DisableDnsParsing", `
                                            "DisableEmailScanning", `
                                            "DisableFtpParsing", `
                                            "DisableGradualRelease", `
                                            "DisableHttpParsing", `
                                            "DisableInboundConnectionFiltering", `
                                            "DisableIOAVProtection", `
                                            "DisableNetworkProtectionPerfTelemetry", `
                                            "DisablePrivacyMode", `
                                            "DisableRdpParsing", `
                                            "DisableRealtimeMonitoring", `
                                            "DisableRemovableDriveScanning", `
                                            "DisableRestorePoint", `
                                            "DisableScanningMappedNetworkDrivesForFullScan", `
                                            "DisableScanningNetworkFiles", `
                                            "DisableScriptScanning", `
                                            "DisableSmtpParsing", `
                                            "DisableSshParsing", `
                                            "DisableTlsParsing", `
                                            "EnableControlledFolderAccess", `
                                            "EnableDnsSinkhole", `
                                            "EnableFileHashComputation", `
                                            "EnableFullScanOnBatteryPower", `
                                            "EnableLowCpuPriority", `
                                            "EnableNetworkProtection", `
                                            "EngineUpdatesChannel", `
                                            "ExclusionExtension", `
                                            "ExclusionIpAddress", `
                                            "ExclusionPath", `
                                            "ExclusionProcess", `
                                            "ForceUseProxyOnly", `
                                            "HighThreatDefaultAction", `
                                            "IntelTDTEnabled", `
                                            "LowThreatDefaultAction", `
                                            "MAPSReporting", `
                                            "MeteredConnectionUpdates", `
                                            "ModerateThreatDefaultAction", `
                                            "PlatformUpdatesChannel", `
                                            "ProxyBypass", `
                                            "ProxyPacUrl", `
                                            "ProxyServer", `
                                            "PSComputerName", `
                                            "PUAProtection", `
                                            "QuarantinePurgeItemsAfterDelay", `
                                            "RandomizeScheduleTaskTimes", `
                                            "RealTimeScanDirection", `
                                            "RemediationScheduleDay", `
                                            "RemediationScheduleTime", `
                                            "ReportDynamicSignatureDroppedEvent", `
                                            "ReportingAdditionalActionTimeOut", `
                                            "ReportingCriticalFailureTimeOut", `
                                            "ReportingNonCriticalTimeOut", `
                                            "ScanAvgCPULoadFactor", `
                                            "ScanOnlyIfIdleEnabled", `
                                            "ScanParameters", `
                                            "ScanPurgeItemsAfterDelay", `
                                            "ScanScheduleDay", `
                                            "ScanScheduleOffset", `
                                            "ScanScheduleQuickScanTime", `
                                            "ScanScheduleTime", `
                                            "SchedulerRandomizationTime", `
                                            "ServiceHealthReportInterval", `
                                            "SevereThreatDefaultAction", `
                                            "SharedSignaturesPath", `
                                            "SignatureAuGracePeriod", `
                                            "SignatureBlobFileSharesSources", `
                                            "SignatureBlobUpdateInterval", `
                                            "SignatureDefinitionUpdateFileSharesSources", `
                                            "SignatureDisableUpdateOnStartupWithoutEngine", `
                                            "SignatureFallbackOrder", `
                                            "SignatureFirstAuGracePeriod", `
                                            "SignatureScheduleDay", `
                                            "SignatureScheduleTime", `
                                            "SignatureUpdateCatchupInterval", `
                                            "SignatureUpdateInterval", `
                                            "SubmitSamplesConsent", `
                                            "ThreatIDDefaultAction_Actions", `
                                            "ThreatIDDefaultAction_Ids", `
                                            "ThrottleForScheduledScanOnly", `
                                            "TrustLabelProtectionStatus", `
                                            "UILockdown", `
                                            "UnknownThreatDefaultAction"
                                            )

                # MPPreference
                    ForEach ($Column in $MPPreferenceArray)
                        {
                            $DefenderObject | add-member -MemberType NoteProperty -Name $Column -Value "" -Force
                        }
            }
    
        # add CollectionTime to existing array
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DefenderObject -Verbose:$Verbose

        # add Computer, ComputerFqdn & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose

    #-------------------------------------------------------------------------------------------
    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
    #-------------------------------------------------------------------------------------------

        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
    #-----------------------------------------------------------------------------------------------
    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
    #-----------------------------------------------------------------------------------------------

        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


###############################################################
# Office [6]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "OFFICE INFORMATION [6]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientOfficeInfoV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"


    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting Office information ... Please Wait !"

        # Default Values
        $OfficeDescription = ""
        $OfficeProductSKU = ""
        $OfficeVersionBuild = ""
        $OfficeInstallationPath = ""
        $OfficeUpdateEnabled = ""
        $OfficeUpdateChannel = ""
        $OfficeUpdateChannelName = ""
        $OneDriveInstalled = ""
        $TeamsInstalled = ""
        $FoundO365Office = $false

        #-----------------------------------------
        # Looking for Microsoft 365 Office
        #-----------------------------------------
            $OfficeVersion  = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
            $OfficePolicies = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate" -ErrorAction SilentlyContinue
                
                #-------------------------------------------------------------------------------------------
                # Preparing data structure
                #-------------------------------------------------------------------------------------------
                If ($OfficeVersion)
                    {
                        $FoundO365Office = $True

                        # Merge results into new OfficeObject
                        $OfficeObject = new-object PSCustomObject
                
                        # $OfficeVersion
                            $ObjColumns = ($OfficeVersion | get-member -MemberType NoteProperty)
                            ForEach ($Entry in $OfficeVersion)
                                {
                                    ForEach ($Column in $ObjColumns)
                                        {
                                            $ColumnName = $Column.name
                                            $OfficeObject | add-member -MemberType NoteProperty -Name $ColumnName -Value $Entry.$ColumnName -force
                                        }
                                }

                        # $OfficePolicies
                            $ObjColumns = ($OfficePolicies | get-member -MemberType NoteProperty)
                            ForEach ($Entry in $OfficePolicies)
                                {
                                    ForEach ($Column in $ObjColumns)
                                        {
                                            $ColumnName = $Column.name
                                            $OfficeObject | add-member -MemberType NoteProperty -Name $ColumnName -Value $Entry.$ColumnName -force
                                        }
                                }

                        $DataVariable = $OfficeObject

                        # convert PS array to PSCustomObject and remove PS class information
                        $DataVariable = Convert-PSArrayToObjectFixStructure -Data $DataVariable -Verbose:$Verbose

                        # add CollectionTime to existing array
                        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                        # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                        # Validating/fixing schema data structure of source data
                        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                        # Aligning data structure with schema (requirement for DCR)
                        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
                    }


        #-----------------------------------------
        # Looking for Office 2016 (standalone)
        #-----------------------------------------
            $OfficeInstallationPath = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot" -Name Path -ErrorAction SilentlyContinue

                #-------------------------------------------------------------------------------------------
                # Preparing data structure
                #-------------------------------------------------------------------------------------------
                If ( ($OfficeInstallationPath) -and ($FoundO365Office -eq $false) )
                    {
                        $OfficeVersionBuild = $Application.Version

                        Switch -Wildcard ($OfficeVersionBuild)
                            {
                                "16.*"    {$OfficeDescription = "Office 2016"}
                            }
                        $OfficeInstallationPath = $OfficeInstallationPath
                        $OfficeUpdateEnabled = $Officeversion.UpdatesChannel
                        $OfficeProductSKU = $OfficeVersion.ProductReleaseIds

                        $DataArray = [pscustomobject]@{
                                                        OfficeDescription       = $OfficeDescription
                                                        OfficeProductSKU        = $OfficeProductSKU
                                                        OfficeVersionBuild      = $OfficeVersionBuild
                                                        OfficeInstallationPath  = $OfficeInstallationPath
                                                        OfficeUpdateEnabled     = $OfficeUpdateEnabled
                                                        OfficeUpdateChannel     = $OfficeUpdateChannel
                                                        OfficeUpdateChannelName = $OfficeUpdateChannelName
                                                      }

                        $DataVariable = $DataArray

                        # add CollectionTime to existing array
                        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                        # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                        # Validating/fixing schema data structure of source data
                        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                        # Aligning data structure with schema (requirement for DCR)
                        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
                    }

        #-----------------------------------------
        # Looking for Office 2013
        #-----------------------------------------

            $OfficeInstallationPath = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Office\15.0\Word\InstallRoot" -Name Path -ErrorAction SilentlyContinue

                #-------------------------------------------------------------------------------------------
                # Preparing data structure
                #-------------------------------------------------------------------------------------------

                If ($OfficeInstallationPath)
                    {
                        $OfficeVersionBuild = $Application.Version

                        Switch -Wildcard ($OfficeVersionBuild)
                            {
                                "15.*"    {$OfficeDescription = "Office 2013"}
                            }

                        $OfficeInstallationPath = $OfficeInstallationPath
                        $OfficeUpdateEnabled = "N/A"
                        $OfficeProductSKU = "N/A"

                        $DataArray = [pscustomobject]@{
                                                        OfficeDescription       = $OfficeDescription
                                                        OfficeProductSKU        = $OfficeProductSKU
                                                        OfficeVersionBuild      = $OfficeVersionBuild
                                                        OfficeInstallationPath  = $OfficeInstallationPath
                                                        OfficeUpdateEnabled     = $OfficeUpdateEnabled
                                                        OfficeUpdateChannel     = $OfficeUpdateChannel
                                                        OfficeUpdateChannelName = $OfficeUpdateChannelName
                                                      }

                        $DataVariable = $DataArray

                        # add CollectionTime to existing array
                        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                        # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                        # Validating/fixing schema data structure of source data
                        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                        # Aligning data structure with schema (requirement for DCR)
                        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
                    }

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                               -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                               -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                               -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                               -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                               -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                             -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


###############################################################
# VPN CLIENT [7]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "VPN INFORMATION [7]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientVpnV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting VPN information ... Please Wait !"

        # Default Values
            $VPNSoftware = ""
            $VPNVersion = ""

            ForEach ($Application in $InstalledApplications)
                {

                    Try
                        {
                            #-----------------------------------------
                            # Looking for Palo Alto
                            #-----------------------------------------
                                If ( ($Application.Vendor -like 'Palo Alto*') -and ($Application.name -like "*Global*") )
                                    {
                                        $VPNSoftware = $Application.Name
                                        $VPNVersion = $Application.Version
                                    }

                                ElseIf ( ($Application.Publisher -like 'Palo Alto*') -and ($Application.name -like "*Global*") )
                                    {
                                        $VPNSoftware = $Application.DisplayName
                                        $VPNVersion = $Application.DisplayVersion
                                    }

                            #-----------------------------------------
                            # Looking for Cisco AnyConnect
                            #-----------------------------------------
                                If ( ($Application.Vendor -like 'Cisco*') -and ($Application.name -like "*AnyConnect*") )
                                    {
                                        $VPNSoftware = $Application.Name
                                        $VPNVersion = $Application.Version
                                    }

                                ElseIf ( ($Application.Publisher -like 'Cisco*') -and ($Application.name -like "*AnyConnect*") )
                                    {
                                        $VPNSoftware = $Application.DisplayName
                                        $VPNVersion = $Application.DisplayVersion
                                    }

                        }
                    Catch
                        {
                        }
                }


    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        $DataVariable = [pscustomobject]@{
                                            VPNSoftware     = $VPNSoftware
                                            VPNVersion      = $VPNVersion
                                         }
            
        # add CollectionTime to existing array
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

        # add Computer, ComputerFqdn & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose

    #-------------------------------------------------------------------------------------------
    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
    #-------------------------------------------------------------------------------------------

        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
    #-----------------------------------------------------------------------------------------------
    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
    #-----------------------------------------------------------------------------------------------

        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


###############################################################
# LAPS [8]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "LAPS INFORMATION [8]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientLAPSInfoV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting LAPS information ... Please Wait !"

        # Default Values
            $LAPSSoftware = ""
            $LAPSVersion = ""

        # Checking
            ForEach ($Application in $InstalledApplications)
                {
                    #-----------------------------------------
                    # Looking for LAPS
                    #-----------------------------------------

                    Try
                        {
                            If ( ($Application.name -like "*Local Administrator Password*") )
                                {
                                    $LAPSSoftware = $Application.Name
                                    $LAPSVersion = $Application.Version
                                }

                            # use alternative name on servers
                            ElseIf ( ($Application.DisplayName -like "*Local Administrator Password*") )
                                {
                                    $LAPSSoftware = $Application.DisplayName
                                    $LAPSVersion = $Application.DisplayVersion
                                }

                        }
                    Catch
                        {
                        }
                }


    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        $DataVariable = [pscustomobject]@{
                                            LAPSSoftware    = $LAPSSoftware
                                            LAPSVersion     = $LAPSVersion
                                         }

        # add CollectionTime to existing array
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

        # add Computer, ComputerFqdn & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose

    #-------------------------------------------------------------------------------------------
    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
    #-------------------------------------------------------------------------------------------

        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
    #-----------------------------------------------------------------------------------------------
    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
    #-----------------------------------------------------------------------------------------------

        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


###############################################################
# ADMIN BY REQUEST [9]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "ADMIN BY REQUEST [9]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientAdminByRequestV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting Admin By Request information ... Please Wait !"

        # Default Values
        $ABRSoftware = ""
        $ABRVersion = ""

        # Checking
            ForEach ($Application in $InstalledApplications)
                {

                    Try
                        {
                            If ( ($Application.name -like "*Admin By Request*") )
                                {
                                    $ABRSoftware = $Application.Name
                                    $ABRVersion = $Application.Version
                                }

                            # use alternative name on servers
                            ElseIf ( ($Application.DisplayName -like "*Admin By Request*") )
                                {
                                    $ABRSoftware = $Application.DisplayName
                                    $ABRVersion = $Application.DisplayVersion
                                }

                        }
                    Catch
                        {
                        }
                }


    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        If ($ABRSoftware)
            {
                $DataVariable = [pscustomobject]@{
                                                    ABRSoftware     = $ABRSoftware
                                                    ABRVersion      = $ABRVersion
                                                 }
    
                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
            }

    #-------------------------------------------------------------------------------------------
    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
    #-------------------------------------------------------------------------------------------

        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
    #-----------------------------------------------------------------------------------------------
    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
    #-----------------------------------------------------------------------------------------------

        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


###############################################################
# Windows Update [10]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-Output "WINDOWS UPDATE INFORMATION [10]"
    Write-Output ""


    #################################################
    # Windows Update Last Results
    #################################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientWindowsUpdateLastResultsV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting Windows Update Last Results information ... Please Wait !"

            $DataVariable = Get-WULastResults

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------

            If ($DataVariable)
                {
                    # convert CIM array to PSCustomObject and remove CIM class information
                    $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable -Verbose:$Verbose
    
                    # add CollectionTime to existing array
                    $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                    # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                    $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                    # Validating/fixing schema data structure of source data
                    $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                    # Aligning data structure with schema (requirement for DCR)
                    $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
                }

    #-------------------------------------------------------------------------------------------
    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
    #-------------------------------------------------------------------------------------------

        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
    #-----------------------------------------------------------------------------------------------
    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
    #-----------------------------------------------------------------------------------------------

        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


    #################################################
    # Windows Update Source Information
    #################################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientWindowsUpdateServiceManagerV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting Windows Update Source information ... Please Wait !"

            $DataVariable = Get-WUServiceManager | Where-Object { $_.IsDefaultAUService -eq $true }

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------

            If ($DataVariable)
                {
                    # convert CIM array to PSCustomObject and remove CIM class information
                    $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable -Verbose:$Verbose
    
                    # add CollectionTime to existing array
                    $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                    # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                    $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                    # Validating/fixing schema data structure of source data
                    $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                    # Aligning data structure with schema (requirement for DCR)
                    $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
                }

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                               -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                               -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                               -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                               -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                               -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                             -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


    #################################################
    # Pending Windows Updates
    #################################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientWindowsUpdatePendingUpdatesV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting information about pending Windows Updates ... Please Wait !"

            $WU_ServiceManager = Get-WUServiceManager | Where-Object { $_.IsDefaultAUService -eq $true }

            # Windows Update
            If ($WU_ServiceManager.ServiceID -eq "9482f4b4-e343-43b6-b170-9a65bc822c77")
                { 
                    $WU_PendingUpdates = Get-WindowsUpdate -WindowsUpdate
                }
            # Microsoft Update
            ElseIf ($WU_ServiceManager.ServiceID -eq "7971f918-a847-4430-9279-4a52d1efe18d")
                { 
                    $WU_PendingUpdates = Get-WindowsUpdate -MicrosoftUpdate
                }
            
            # WSUS
            Elseif ($WU_ServiceManager.ServiceID -eq "3da21691-e39d-4da6-8a4b-b43877bcb1b7")
                { 
                    $WU_PendingUpdates = Get-WindowsUpdate -ServiceID $WU_ServiceManager.ServiceID
                }

            # other
            Else
                { 
                    $WU_PendingUpdates = Get-WindowsUpdate
                }

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------

            If ($WU_PendingUpdates)
                {
                    $CountDataVariable = ($WU_PendingUpdates | Measure-Object).count

                    $PosDataVariable   = 0
                    Do
                        {
                            # CVEs
                                $UpdateCVEs = $WU_PendingUpdates[$PosDataVariable].CveIDs -join ";"
                                $WU_PendingUpdates[$PosDataVariable] | Add-Member -Type NoteProperty -Name 'UpdateCVEs' -Value $UpdateCVEs -force

                            # Classification (e.g. Security Update)

                                $UpdateClassificationCount = ($WU_PendingUpdates[$PosDataVariable].Categories | Measure-Object).count

                                $UpdClassification = ""
                                $UpdTarget         = ""
                                ForEach ($Classification in $WU_PendingUpdates[$PosDataVariable].Categories)
                                    {
                                        
                                        If ($Classification.Type -eq "UpdateClassification")
                                            {
                                                $UpdClassification = $Classification.name
                                            }
                                        ElseIf ($Classification.Type -ne "UpdateClassification")
                                            {
                                                $UpdTarget = $Classification.name
                                            }
                                    }

                                $WU_PendingUpdates[$PosDataVariable] | Add-Member -Type NoteProperty -Name 'UpdateClassification' -Value $UpdClassification -force
                                $WU_PendingUpdates[$PosDataVariable] | Add-Member -Type NoteProperty -Name 'UpdateTarget' -Value $UpdTarget -force

                            # Target (e.g. product, SQL)
                                $UpdateTarget     = $WU_PendingUpdates[$PosDataVariable].Categories | Where-Object { $_.Type -ne "UpdateClassification" } | Select Name
                                $UpdateTargetName = $UpdateTarget.Name
                                $WU_PendingUpdates[$PosDataVariable] | Add-Member -Type NoteProperty -Name 'UpdateTarget' -Value $UpdateTargetName -force

                            # KB
                                $UpdateKB = $WU_PendingUpdates[$PosDataVariable].KBArticleIDs -join ";"
                                $WU_PendingUpdates[$PosDataVariable] | Add-Member -Type NoteProperty -Name 'UpdateKB' -Value $UpdateKB -force

                            # KB Published Date
                                $UpdateKBPublished                 = $WU_PendingUpdates[$PosDataVariable].LastDeploymentChangeTime
                                $WU_PendingUpdates[$PosDataVariable] | Add-Member -Type NoteProperty -Name 'UpdateKBPublished' -Value $UpdateKBPublished -force

                            $PosDataVariable = 1 + $PosDataVariable
                        }
                    Until ($PosDataVariable -eq $CountDataVariable)


                    # add CollectionTime to existing array
                    $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $WU_PendingUpdates -Verbose:$Verbose

                    # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                    $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                    # Remove unnecessary columns in schema
                    $DataVariable = Filter-ObjectExcludeProperty -Data $DataVariable -ExcludeProperty BundledUpdates, Categories, CveIds, DownloadContents, Identity, InstallationBehavior, UninstallationSteps, `
                                                                                                      KBArticleIDs, Languages, MoreInfoUrls, SecurityBulletinIDs, SuperSededUpdateIds, UninstallationBehavior, WindowsDriverUpdateEntries -Verbose:$Verbose

                    # Validating/fixing schema data structure of source data
                    $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                    # Aligning data structure with schema (requirement for DCR)
                    $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose


                #-------------------------------------------------------------------------------------------
                # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
                #-------------------------------------------------------------------------------------------

                    $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                                       -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                                       -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                                       -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                                       -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                                       -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                                       -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
                #-----------------------------------------------------------------------------------------------
                # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
                #-----------------------------------------------------------------------------------------------

                    $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                                     -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose
                }


    #############################################################
    # Status of Windows Update installations
    #############################################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientWindowsUpdateLastInstallationsV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting information about installations of Windows Updates (incl. A/V updates) ... Please Wait !"

            $OsInfo = Get-CimInstance -ClassName Win32_OperatingSystem
            $ProductType = $OsInfo.ProductType  # 1 = workstation, 2 = domain controller, 3 = server

            # Collection (servers)
            If ( ($ProductType -eq "2") -or ($ProductType -eq "3") )
                {
                    # Getting OS install-date
                    $DaysSinceInstallDate = (Get-Date) - (Get-date $OSInfo.InstallDate)

                    If ([version]$OSInfo.Version -gt "6.3.9600")  # Win2016 and higher
                        { 
                            Write-Verbose "Win2016 or higher detected (last 1000 updates incl. A/V updates)"
                            $Installed_Updates_PSWindowsUpdate_All = Get-WUHistory -MaxDate $DaysSinceInstallDate.Days -Last 1000
                        }
                    ElseIf ([version]$OSInfo.Version -le "6.3.9600")  # Win2012 R2 or Win2012
                        {
                            Write-Verbose "Windows2012 / Win2012 R2 detected (last 100 updates incl. A/V updates)"
                            $Installed_Updates_PSWindowsUpdate_All = Get-WUHistory -Last 100
                        }
                    Else
                        {
                            Write-Verbose "No collection of installed updates"
                            $Installed_Updates_PSWindowsUpdate_All = ""
                        }
                }

            # Collection (workstations)
            If ($ProductType -eq "1")
                {
                    # Getting OS install-date
                    $DaysSinceInstallDate = (Get-Date) - (Get-date $OSInfo.InstallDate)

                    $Installed_Updates_PSWindowsUpdate_All = Get-WUHistory -MaxDate $DaysSinceInstallDate.Days -Last 1000
                }

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------

            If ($Installed_Updates_PSWindowsUpdate_All)
                {
                    Write-Verbose "Processing Windows updates object"
                    $CountDataVariable = ($Installed_Updates_PSWindowsUpdate_All | Measure-Object).count

                    $PosDataVariable   = 0
                    Do
                        {
                            # CVEs
                                $UpdateCVEsInfo = $Installed_Updates_PSWindowsUpdate_All[$PosDataVariable].CveIDs -join ";"
                                $Installed_Updates_PSWindowsUpdate_All[$PosDataVariable] | Add-Member -Type NoteProperty -Name 'UpdateCVEs' -Value $UpdateCVEsInfo -force

                            # Classification (e.g. Security Update)

                                $UpdClassification = ""
                                $UpdTarget = ""
                                ForEach ($Classification in $Installed_Updates_PSWindowsUpdate_All[$PosDataVariable].Categories)
                                    {
                                        
                                        If ($Classification.Type -eq "UpdateClassification")
                                            {
                                                $UpdClassification = $Classification.name
                                            }
                                        ElseIf ($Classification.Type -eq "Product")
                                            {
                                                $UpdTarget = $Classification.type
                                            }
                                    }

                                $Installed_Updates_PSWindowsUpdate_All[$PosDataVariable] | Add-Member -Type NoteProperty -Name 'UpdateClassification' -Value $UpdClassification -force
                                $Installed_Updates_PSWindowsUpdate_All[$PosDataVariable] | Add-Member -Type NoteProperty -Name 'UpdateTarget' -Value $UpdTarget -force

                            # KB
                                $KB = ($Installed_Updates_PSWindowsUpdate_All[$PosDataVariable].KBArticleIDs -join ";")
                                If ($KB)
                                    {
                                        $UpdateKB = "KB" + $KB
                                    }
                                $Installed_Updates_PSWindowsUpdate_All[$PosDataVariable] | Add-Member -Type NoteProperty -Name 'UpdateKB' -Value $UpdateKB -force

                            # KB Published Date
                                $UpdateKBPublished = $Installed_Updates_PSWindowsUpdate_All[$PosDataVariable].LastDeploymentChangeTime
                                $Installed_Updates_PSWindowsUpdate_All[$PosDataVariable] | Add-Member -Type NoteProperty -Name 'UpdateKBPublished' -Value $UpdateKBPublished -force

                            $PosDataVariable = 1 + $PosDataVariable
                        }
                    Until ($PosDataVariable -eq $CountDataVariable)

                    # Remove unnecessary columns in schema
                    $DataVariable = Filter-ObjectExcludeProperty -Data $Installed_Updates_PSWindowsUpdate_All -ExcludeProperty UninstallationSteps,Categories,UpdateIdentity,UnMappedResultCode,UninstallationNotes,HResult -Verbose:$Verbose

                    # convert CIM array to PSCustomObject and remove CIM class information
                    $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable -Verbose:$Verbose

                    # add CollectionTime to existing array
                    $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                    # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                    $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                    # Validating/fixing schema data structure of source data
                    $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                    # Aligning data structure with schema (requirement for DCR)
                    $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
            }
        Else
            {
                $DataVariable = ""
            }


    #-------------------------------------------------------------------------------------------
    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
    #-------------------------------------------------------------------------------------------

        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
    #-----------------------------------------------------------------------------------------------
    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
    #-----------------------------------------------------------------------------------------------

        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


###############################################################
# Bitlocker [11]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "BITLOCKER INFORMATION [11]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientBitlockerInfoV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting information about Bitlocker ... Please Wait !"

        # Default Values
        $OSDisk_DriveLetter = ""
        $OSDisk_CapacityGB = ""
        $OSDisk_VolumeStatus = ""
        $OSDisk_EncryptionPercentage = ""
        $OSDisk_KeyProtector = ""
        $OSDisk_AutoUnlockEnabled = ""
        $OSDisk_ProtectionStatus = ""

        # Step 1/3 - get information
        Try
            {
                $BitlockerVolumens = Get-BitLockerVolume 
            }
        Catch
            {
                Write-output "  Bitlocker was not found on this machine !!"
            }


        If ($BitlockerVolumens)
            {
                # OS Disk
                $OSVolumen = $BitLockerVolumens | where VolumeType -EQ "OperatingSystem"
                $OSDisk_DriveLetter = $OSVOlumen.MountPoint
                $OSDisk_CapacityGB = $OSVOlumen.CapacityGB
                $OSDisk_VolumeStatus = $OSVOlumen.VolumeStatus
                $OSDisk_EncryptionPercentage = $OSVOlumen.EncryptionPercentage
                $OSDisk_KeyProtector = $OSVOlumen.KeyProtector
                $OSDisk_AutoUnlockEnabled = $OSVOlumen.AutoUnlockEnabled
                $OSDisk_ProtectionStatus = $OSVOlumen.ProtectionStatus
            }

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        $DataVariable = [pscustomobject]@{
                                            OSDisk_DriveLetter = $OSDisk_DriveLetter
                                            OSDisk_CapacityGB= $OSDisk_CapacityGB
                                            OSDisk_VolumeStatus = $OSDisk_VolumeStatus
                                            OSDisk_EncryptionPercentage = $OSDisk_EncryptionPercentage
                                            OSDisk_KeyProtector = $OSDisk_KeyProtector
                                            OSDisk_AutoUnlockEnabled = $OSDisk_AutoUnlockEnabled
                                            OSDisk_ProtectionStatus = $OSDisk_ProtectionStatus
                                         }

        # convert CIM array to PSCustomObject and remove CIM class information
        $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable -Verbose:$Verbose
    
        # add CollectionTime to existing array
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

        # add Computer, ComputerFqdn & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose

    #-------------------------------------------------------------------------------------------
    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
    #-------------------------------------------------------------------------------------------

        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
    #-----------------------------------------------------------------------------------------------
    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
    #-----------------------------------------------------------------------------------------------

        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


###############################################################
# EVENTLOG [12]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "EVENTLOG [12]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientEventlogInfoV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting Eventlog information ... Please Wait !"

        $FilteredEvents      = @()
        $Appl_Events_ALL     = @()
        $System_Events_ALL   = @()
        $Security_Events_ALL = @()
        $DataVariable        = @()

        ###############################################################################################

        $Application_EventId_Array = @(
                                      )


        $System_EventId_Array      = @(
                                        "6008;Eventlog"  # Unexpected shutdown ; Providername = Eventlog
                                        "7001;Microsoft-Windows-WinLogon" # Windows logon
                                      )

        $Security_EventId_Array    = @( 
                                        "4740;Microsoft-Windows-Security-Auditing"  # Accounts Lockouts
                                        "4728;Microsoft-Windows-Security-Auditing"  # User Added to Privileged Group
                                        "4732;Microsoft-Windows-Security-Auditing"  # User Added to Privileged Group
                                        "4756;Microsoft-Windows-Security-Auditing"  # User Added to Privileged Group
                                        "4735;Microsoft-Windows-Security-Auditing"  # Security-Enabled Group Modification
                                        "4625;Microsoft-Windows-Security-Auditing"  # Failed User Account Login
                                        "4648;Microsoft-Windows-Security-Auditing"  # Account Login with Explicit Credentials
                                      )

        ###############################################################################################

        $Yesterday = (Get-Date).AddDays(-1)
            
        If ($Application_EventId_Array)
            {
                    ForEach ($Entry in $Application_EventId_Array)
                        {
                            $Split = $Entry -split ";"
                            $Id    = $Split[0]
                            $ProviderName = $Split[1]

                            $FilteredEvents += Get-WinEvent -FilterHashtable @{ProviderName = $ProviderName; Id = $Id} -ErrorAction SilentlyContinue | Where-Object { ($_.TimeCreated -ge $Yesterday) }
                        }
            }

        If ($System_EventId_Array)
            {
                    ForEach ($Entry in $System_EventId_Array)
                        {
                            $Split = $Entry -split ";"
                            $Id    = $Split[0]
                            $ProviderName = $Split[1]

                            $FilteredEvents += Get-WinEvent -FilterHashtable @{ProviderName = $ProviderName; Id = $Id} -ErrorAction SilentlyContinue | Where-Object { ($_.TimeCreated -ge $Yesterday) }
                        }
            }

        If ($Security_EventId_Array)
            {
                    ForEach ($Entry in $Security_EventId_Array)
                        {
                            $Split = $Entry -split ";"
                            $Id    = $Split[0]
                            $ProviderName = $Split[1]

                            $FilteredEvents += Get-WinEvent -FilterHashtable @{ProviderName = $ProviderName; Id = $Id} -ErrorAction SilentlyContinue | Where-Object { ($_.TimeCreated -ge $Yesterday) }
                        }
            }

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        If ($FilteredEvents)
            {
                # convert CIM array to PSCustomObject and remove CIM class information
                $DataVariable = Convert-CimArrayToObjectFixStructure -data $FilteredEvents -Verbose:$Verbose
    
                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose

            #-------------------------------------------------------------------------------------------
            # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
            #-------------------------------------------------------------------------------------------

                $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                                   -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                                   -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                                   -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                                   -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                                   -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                                   -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
            #-----------------------------------------------------------------------------------------------
            # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
            #-----------------------------------------------------------------------------------------------

                $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose
            }


###############################################################
# Network Adapter Information [13]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "NETWORK ADAPTER INFORMATION [13]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientNetworkAdapterInfoV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting Network Adapter information ... Please Wait !"

        $NetworkAdapter = Get-NetAdapter

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------
        If ($NetworkAdapter)
            {
                # convert CIM array to PSCustomObject and remove CIM class information
                $DataVariable = Convert-CimArrayToObjectFixStructure -data $NetworkAdapter -Verbose:$Verbose
    
                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                # Get insight about the schema structure of an object BEFORE changes. Command is only needed to verify schema - can be disabled
                $SchemaBefore = Get-ObjectSchemaAsArray -Data $DataVariable -Verbose:$Verbose
        
                # Remove unnecessary columns in schema
                $DataVariable = Filter-ObjectExcludeProperty -Data $DataVariable -ExcludeProperty Memento*,Inno*,'(default)',1033 -Verbose:$Verbose

                # Get insight about the schema structure of an object AFTER changes. Command is only needed to verify schema - can be disabled
                $Schema = Get-ObjectSchemaAsArray -Data $DataVariable -Verbose:$Verbose

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose            }
        Else
            {

                # log issue - typically WMI issue
                $TableName  = 'InvClientCollectionIssuesV2'   # must not contain _CL
                $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

                $DataVariable = [pscustomobject]@{
                                                   IssueCategory   = "NetworkAdapterInformation"
                                                 }

                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
            }         

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                               -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                               -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                               -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                               -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                               -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                             -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


###############################################################
# IP INFORMATION [14]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "IP INFORMATION [14]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientNetworkIPv4InfoV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting IPv4 information ... Please Wait !"

        $IPv4Status = Get-NetIPAddress -AddressFamily IPv4

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------
        If ($IPv4Status)
            {
                # convert CIM array to PSCustomObject and remove CIM class information
                $DataVariable = Convert-CimArrayToObjectFixStructure -data $IPv4Status -Verbose:$Verbose
    
                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
            }
        Else
            {

                # log issue - typically WMI issue
                $TableName  = 'InvClientCollectionIssuesV2'   # must not contain _CL
                $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

                $DataVariable = [pscustomobject]@{
                                                   IssueCategory   = "IPInformation"
                                                 }

                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName  -Column2Name UserLoggedOn -Column2Data $UserLoggedOn -Verbose:$Verbose

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
            }         

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                               -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                               -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                               -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                               -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                               -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                             -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


###############################################################
# LOCAL ADMINISTRATORS GROUP [15]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "LOCAL ADMINISTRATORS GROUP INFORMATION [15]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientLocalAdminsV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting Local Admin information ... Please Wait !"

        $LocalAdminGroup = (Get-localgroup -Sid S-1-5-32-544).name       # SID S-1-5-32-544 = local computer Administrators group

        $DataVariable = @()
        foreach ($Object in Get-LocalGroup -Name $LocalAdminGroup ) 
            {
              $MemberChk = [ADSI]"WinNT://$env:COMPUTERNAME/$Object"
              $MemberList = @($MemberChk.Invoke('Members') | % {([adsi]$_).path})
      
              ForEach ($MemberEntry in $MemberList)
                {
                  # Strip type into separate column
                  $type = $MemberEntry.substring(0,8)
                  $member = $MemberEntry.substring(8)
          
                  If ($member.Contains("/"))  # exclude SID
                    {
                      $LocalAdminObj = New-object PsCustomObject
                      $LocalAdminObj | Add-Member -MemberType NoteProperty -Name "Name" -Value $member
                      $LocalAdminObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type
                      $DataVariable += $LocalAdminObj
                    }
                }
            }


    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        If ($DataVariable)
            {
                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
			 

                #-------------------------------------------------------------------------------------------
                # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
                #-------------------------------------------------------------------------------------------

                    $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                                       -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                                       -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                                       -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                                       -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                                       -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                                       -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
                #-----------------------------------------------------------------------------------------------
                # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
                #-----------------------------------------------------------------------------------------------

                    $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                                     -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose
            }

#####################################################################
# WINDOWS FIREWALL [16]
#####################################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "WINDOWS FIREWALL INFORMATION [16]"
    Write-output ""

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientWindowsFirewallInfoV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting Windows Firewall information ... Please Wait !"

            $WinFw = Get-NetFirewallProfile -policystore activestore

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------
        If ($WinFw)
            {
                # convert CIM array to PSCustomObject and remove CIM class information
                $DataVariable = Convert-CimArrayToObjectFixStructure -data $WinFw -Verbose:$Verbose
    
                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
            }
        Else
            {

                # log issue - typically WMI issue
                $TableName  = 'InvClientCollectionIssuesV2'   # must not contain _CL
                $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

                $DataVariable = [pscustomobject]@{
                                                   IssueCategory   = "WinFwInformation"
                                                 }

                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName  -Column2Name UserLoggedOn -Column2Data $UserLoggedOn -Verbose:$Verbose

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
            }         

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                               -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                               -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                               -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                               -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                               -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                             -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose

###############################################################
# GROUP POLICY [17]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "GROUP POLICY INFORMATION [17]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientGroupPolicyRefreshV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting Group Policy information ... Please Wait !"

        # Get StartTimeHi Int32 value
        $startTimeHi = (Get-ItemProperty -Path "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Extension-List\{00000000-0000-0000-0000-000000000000}").startTimeHi
            
        # Get StartTimeLo Int32 value
        $startTimeLo = (Get-ItemProperty -Path "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Extension-List\{00000000-0000-0000-0000-000000000000}").startTimeLo
            
        $GPLastRefresh = [datetime]::FromFileTime(([Int64] ((Get-ItemProperty -Path "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Extension-List\{00000000-0000-0000-0000-000000000000}").startTimeHi) -shl 32) -bor ((Get-ItemProperty -Path "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Extension-List\{00000000-0000-0000-0000-000000000000}").startTimeLo))
        $CalculateGPLastRefreshTimeSpan = NEW-TIMESPAN –Start $GPLastRefresh –End (Get-date)

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        If ($GPLastRefresh)
            {
                $DataArray = [pscustomobject]@{
                                                GPLastRefresh       = $GPLastRefresh
                                                GPLastRefreshDays   = $CalculateGPLastRefreshTimeSpan.Days
                                              }
   
                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataArray -Verbose:$Verbose

                # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
            }

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                               -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                               -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                               -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                               -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                               -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                             -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


###############################################################
# TPM [18]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "TPM [18]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientHardwareTPMInfoV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefix + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output ""
        Write-Output "Collecting TPM information ... Please Wait !"

        $TPM = Get-TPM -ErrorAction SilentlyContinue -WarningVariable SilentlyContinue

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        If ($TPM)
            {

                # convert CIM array to PSCustomObject and remove CIM class information
                $TPM = Convert-CimArrayToObjectFixStructure -data $TPM -Verbose:$Verbose

                # Get TPM Version, cannot be found using Get-TPM - must be retrieved from WMI
                $TPMInfo_WMI = Get-CimInstance -Namespace "Root\CIMV2\Security\MicrosoftTpm" -query "Select * from Win32_Tpm"
                If ($TPMInfo_WMI)
                    {
                        $TPM_Version_WMI_Major = (($TPMInfo_WMI.SpecVersion.split(","))[0])
                        $TPM_Version_WMI_Major = $TPM_Version_WMI_Major.trim()

                        $TPM_Version_WMI_Minor = (($TPMInfo_WMI.SpecVersion.split(","))[1])
                        $TPM_Version_WMI_Minor = $TPM_Version_WMI_Minor.trim()

                        $TPM_Version_WMI_Rev = (($TPMInfo_WMI.SpecVersion.split(","))[2])
                        $TPM_Version_WMI_Rev = $TPM_Version_WMI_Rev.trim()
                    }

                $TPMCount = 0
                ForEach ($Entry in $TPM)
                    {
                        $TPMCount = 1 + $TPMCount
                    }

                $CountDataVariable = $TPMCount
                $PosDataVariable   = 0
                Do
                    {
                        $TPM[$PosDataVariable] | Add-Member -Type NoteProperty -Name TPM_Version_WMI_Major -Value $TPM_Version_WMI_Major -force
                        $TPM[$PosDataVariable] | Add-Member -Type NoteProperty -Name TPM_Version_WMI_Minor -Value $TPM_Version_WMI_Minor -force
                        $TPM[$PosDataVariable] | Add-Member -Type NoteProperty -Name TPM_Version_WMI_Rev -Value $TPM_Version_WMI_Rev -force
                        $PosDataVariable = 1 + $PosDataVariable
                    }
                Until ($PosDataVariable -eq $CountDataVariable)

                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $TPM -Verbose:$Verbose

                # add Computer, ComputerFqdn & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name ComputerFqdn -Column2Data $DnsName -Column3Name UserLoggedOn -Column3Data $UserLoggedOn -Verbose:$Verbose

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose
            }
        Else
            {
                $DataVariable = [pscustomobject]@{
                                                    IssueCategory   = "TPM"
                                                 }

                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$Verbose

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName  -Column2Name UserLoggedOn -Column2Data $UserLoggedOn -Verbose:$Verbose

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$Verbose

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$Verbose
            }

    #-------------------------------------------------------------------------------------------
    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
    #-------------------------------------------------------------------------------------------

        $ResultMgmt = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose `
                                                           -DceName $DceName -DcrName $DcrName -DcrResourceGroup $AzDcrResourceGroup -TableName $TableName -Data $DataVariable `
                                                           -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                           -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                           -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                           -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
    #-----------------------------------------------------------------------------------------------
    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
    #-----------------------------------------------------------------------------------------------

        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable -TableName $TableName `
                                                                         -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId -Verbose:$Verbose


##################################
# WRITE LASTRUN KEY
##################################

    $Now = (Get-date)

    # Create initial reg-path stucture in registry
        If (-not (Test-Path $LastRun_RegPath))
            {
                $Err = New-Item -Path $LastRun_RegPath -Force | Out-Null
            }

    #  Set last run value in registry
        $Result = New-ItemProperty -Path $LastRun_RegPath -Name $LastRun_RegKey -Value $Now -PropertyType STRING -Force | Out-Null


# SIG # Begin signature block
# MIIRgwYJKoZIhvcNAQcCoIIRdDCCEXACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQULRP1XZscf2HTwVkxSdZC2m3e
# lPSggg3jMIIG5jCCBM6gAwIBAgIQd70OA6G3CPhUqwZyENkERzANBgkqhkiG9w0B
# AQsFADBTMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2lnbiBudi1zYTEp
# MCcGA1UEAxMgR2xvYmFsU2lnbiBDb2RlIFNpZ25pbmcgUm9vdCBSNDUwHhcNMjAw
# NzI4MDAwMDAwWhcNMzAwNzI4MDAwMDAwWjBZMQswCQYDVQQGEwJCRTEZMBcGA1UE
# ChMQR2xvYmFsU2lnbiBudi1zYTEvMC0GA1UEAxMmR2xvYmFsU2lnbiBHQ0MgUjQ1
# IENvZGVTaWduaW5nIENBIDIwMjAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIK
# AoICAQDWQk3540/GI/RsHYGmMPdIPc/Q5Y3lICKWB0Q1XQbPDx1wYOYmVPpTI2AC
# qF8CAveOyW49qXgFvY71TxkkmXzPERabH3tr0qN7aGV3q9ixLD/TcgYyXFusUGcs
# JU1WBjb8wWJMfX2GFpWaXVS6UNCwf6JEGenWbmw+E8KfEdRfNFtRaDFjCvhb0N66
# WV8xr4loOEA+COhTZ05jtiGO792NhUFVnhy8N9yVoMRxpx8bpUluCiBZfomjWBWX
# ACVp397CalBlTlP7a6GfGB6KDl9UXr3gW8/yDATS3gihECb3svN6LsKOlsE/zqXa
# 9FkojDdloTGWC46kdncVSYRmgiXnQwp3UrGZUUL/obLdnNLcGNnBhqlAHUGXYoa8
# qP+ix2MXBv1mejaUASCJeB+Q9HupUk5qT1QGKoCvnsdQQvplCuMB9LFurA6o44EZ
# qDjIngMohqR0p0eVfnJaKnsVahzEaeawvkAZmcvSfVVOIpwQ4KFbw7MueovE3vFL
# H4woeTBFf2wTtj0s/y1KiirsKA8tytScmIpKbVo2LC/fusviQUoIdxiIrTVhlBLz
# pHLr7jaep1EnkTz3ohrM/Ifll+FRh2npIsyDwLcPRWwH4UNP1IxKzs9jsbWkEHr5
# DQwosGs0/iFoJ2/s+PomhFt1Qs2JJnlZnWurY3FikCUNCCDx/wIDAQABo4IBrjCC
# AaowDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMBIGA1UdEwEB
# /wQIMAYBAf8CAQAwHQYDVR0OBBYEFNqzjcAkkKNrd9MMoFndIWdkdgt4MB8GA1Ud
# IwQYMBaAFB8Av0aACvx4ObeltEPZVlC7zpY7MIGTBggrBgEFBQcBAQSBhjCBgzA5
# BggrBgEFBQcwAYYtaHR0cDovL29jc3AuZ2xvYmFsc2lnbi5jb20vY29kZXNpZ25p
# bmdyb290cjQ1MEYGCCsGAQUFBzAChjpodHRwOi8vc2VjdXJlLmdsb2JhbHNpZ24u
# Y29tL2NhY2VydC9jb2Rlc2lnbmluZ3Jvb3RyNDUuY3J0MEEGA1UdHwQ6MDgwNqA0
# oDKGMGh0dHA6Ly9jcmwuZ2xvYmFsc2lnbi5jb20vY29kZXNpZ25pbmdyb290cjQ1
# LmNybDBWBgNVHSAETzBNMEEGCSsGAQQBoDIBMjA0MDIGCCsGAQUFBwIBFiZodHRw
# czovL3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0b3J5LzAIBgZngQwBBAEwDQYJ
# KoZIhvcNAQELBQADggIBAAiIcibGr/qsXwbAqoyQ2tCywKKX/24TMhZU/T70MBGf
# j5j5m1Ld8qIW7tl4laaafGG4BLX468v0YREz9mUltxFCi9hpbsf/lbSBQ6l+rr+C
# 1k3MEaODcWoQXhkFp+dsf1b0qFzDTgmtWWu4+X6lLrj83g7CoPuwBNQTG8cnqbmq
# LTE7z0ZMnetM7LwunPGHo384aV9BQGf2U33qQe+OPfup1BE4Rt886/bNIr0TzfDh
# 5uUzoL485HjVG8wg8jBzsCIc9oTWm1wAAuEoUkv/EktA6u6wGgYGnoTm5/DbhEb7
# c9krQrbJVzTHFsCm6yG5qg73/tvK67wXy7hn6+M+T9uplIZkVckJCsDZBHFKEUta
# ZMO8eHitTEcmZQeZ1c02YKEzU7P2eyrViUA8caWr+JlZ/eObkkvdBb0LDHgGK89T
# 2L0SmlsnhoU/kb7geIBzVN+nHWcrarauTYmAJAhScFDzAf9Eri+a4OFJCOHhW9c4
# 0Z4Kip2UJ5vKo7nb4jZq42+5WGLgNng2AfrBp4l6JlOjXLvSsuuKy2MIL/4e81Yp
# 4jWb2P/ppb1tS1ksiSwvUru1KZDaQ0e8ct282b+Awdywq7RLHVg2N2Trm+GFF5op
# ov3mCNKS/6D4fOHpp9Ewjl8mUCvHouKXd4rv2E0+JuuZQGDzPGcMtghyKTVTgTTc
# MIIG9TCCBN2gAwIBAgIMeWPZY2rjO3HZBQJuMA0GCSqGSIb3DQEBCwUAMFkxCzAJ
# BgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMS8wLQYDVQQDEyZH
# bG9iYWxTaWduIEdDQyBSNDUgQ29kZVNpZ25pbmcgQ0EgMjAyMDAeFw0yMzAzMjcx
# MDIxMzRaFw0yNjAzMjMxNjE4MThaMGMxCzAJBgNVBAYTAkRLMRAwDgYDVQQHEwdL
# b2xkaW5nMRAwDgYDVQQKEwcybGlua0lUMRAwDgYDVQQDEwcybGlua0lUMR4wHAYJ
# KoZIhvcNAQkBFg9tb2tAMmxpbmtpdC5uZXQwggIiMA0GCSqGSIb3DQEBAQUAA4IC
# DwAwggIKAoICAQDMpI1rTOoWOSET3lSFQfsl/t83DCUEdoI02fNS5xlURPeGZNhi
# xQMKrhmFrdbIaEx01eY+hH9gF2AQ1ZDa7orCVSde1LDBnbFPLqcHWW5RWyzcy8Pq
# gV1QvzlFbmvTNHLm+wn1DZJ/1qJ+A+4uNUMrg13WRTiH0YWd6pwmAiQkoGC6FFwE
# usXotrT5JJNcPGlxBccm8su3kakI5B6iEuTeKh92EJM/km0pc/8o+pg+uR+f07Pp
# WcV9sS//JYCSLaXWicfrWq6a7/7U/vp/Wtdz+d2DcwljpsoXd++vuwzF8cUs09uJ
# KtdyrN8Z1DxqFlMdlD0ZyR401qAX4GO2XdzH363TtEBKAwvV+ReW6IeqGp5FUjnU
# j0RZ7NPOSiPr5G7d23RutjCHlGzbUr+5mQV/IHGL9LM5aNHsu22ziVqImRU9nwfq
# QVb8Q4aWD9P92hb3jNcH4bIWiQYccf9hgrMGGARx+wd/vI+AU/DfEtN9KuLJ8rNk
# LfbXRSB70le5SMP8qK09VjNXK/i6qO+Hkfh4vfNnW9JOvKdgRnQjmNEIYWjasbn8
# GyvoFVq0GOexiF/9XFKwbdGpDLJYttfcVZlBoSMPOWRe8HEKZYbJW1McjVIpWPnP
# d6tW7CBY2jp4476OeoPpMiiApuc7BhUC0VWl1Ei2PovDUoh/H3euHrWqbQIDAQAB
# o4IBsTCCAa0wDgYDVR0PAQH/BAQDAgeAMIGbBggrBgEFBQcBAQSBjjCBizBKBggr
# BgEFBQcwAoY+aHR0cDovL3NlY3VyZS5nbG9iYWxzaWduLmNvbS9jYWNlcnQvZ3Nn
# Y2NyNDVjb2Rlc2lnbmNhMjAyMC5jcnQwPQYIKwYBBQUHMAGGMWh0dHA6Ly9vY3Nw
# Lmdsb2JhbHNpZ24uY29tL2dzZ2NjcjQ1Y29kZXNpZ25jYTIwMjAwVgYDVR0gBE8w
# TTBBBgkrBgEEAaAyATIwNDAyBggrBgEFBQcCARYmaHR0cHM6Ly93d3cuZ2xvYmFs
# c2lnbi5jb20vcmVwb3NpdG9yeS8wCAYGZ4EMAQQBMAkGA1UdEwQCMAAwRQYDVR0f
# BD4wPDA6oDigNoY0aHR0cDovL2NybC5nbG9iYWxzaWduLmNvbS9nc2djY3I0NWNv
# ZGVzaWduY2EyMDIwLmNybDATBgNVHSUEDDAKBggrBgEFBQcDAzAfBgNVHSMEGDAW
# gBTas43AJJCja3fTDKBZ3SFnZHYLeDAdBgNVHQ4EFgQUMcaWNqucqymu1RTg02YU
# 3zypsskwDQYJKoZIhvcNAQELBQADggIBAHt/DYGUeCFfbtuuP5/44lpR2wbvOO49
# b6TenaL8TL3VEGe/NHh9yc3LxvH6PdbjtYgyGZLEooIgfnfEo+WL4fqF5X2BH34y
# EAsHCJVjXIjs1mGc5fajx14HU52iLiQOXEfOOk3qUC1TF3NWG+9mezho5XZkSMRo
# 0Ypg7Js2Pk3U7teZReCJFI9FSYa/BT2DnRFWVTlx7T5lIz6rKvTO1qQC2G3NKVGs
# HMtBTjsF6s2gpOzt7zF3o+DsnJukQRn0R9yTzgrx9nXYiHz6ti3HuJ4U7i7ILpgS
# RNrzmpVXXSH0wYxPT6TLm9eZR8qdZn1tGSb1zoIT70arnzE90oz0x7ej1fC8IUA/
# AYhkmfa6feI7OMU5xnsUjhSiyzMVhD06+RD3t5JrbKRoCgqixGb7DGM+yZVjbmhw
# cvr3UGVld9++pbsFeCB3xk/tcMXtBPdHTESPvUjSCpFbyldxVLU6GVIdzaeHAiBy
# S0NXrJVxcyCWusK41bJ1jP9zsnnaUCRERjWF5VZsXYBhY62NSOlFiCNGNYmVt7fi
# b4V6LFGoWvIv2EsWgx/uR/ypWndjmV6uBIN/UMZAhC25iZklNLFGDZ5dCUxLuoyW
# PVCTBYpM3+bN6dmbincjG0YDeRjTVfPN5niP1+SlRwSQxtXqYoDHq+3xVzFWVBqC
# NdoiM/4DqJUBMYIDCjCCAwYCAQEwaTBZMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQ
# R2xvYmFsU2lnbiBudi1zYTEvMC0GA1UEAxMmR2xvYmFsU2lnbiBHQ0MgUjQ1IENv
# ZGVTaWduaW5nIENBIDIwMjACDHlj2WNq4ztx2QUCbjAJBgUrDgMCGgUAoHgwGAYK
# KwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIB
# BDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU
# DWZ+TS4yW2PMyX09+PLvHymbFWUwDQYJKoZIhvcNAQEBBQAEggIAaNFSOcleqdGg
# uXutLWUyhEs0/QgIG6tTJHFPbEwvi+4xG6HmmH00Npsd8pzvN7726vjUMbzB0WPx
# aMfQAuRFxSRrundHSiU941VLZ/7L2IFpav6fpa/R2YtQmokgZK0xJ6zyKkTd/xd/
# GWWx0BcANpDBoNWoaa6cwHIh7rqg6XkQiGzX5XuRd14CYQm+exkw8A3HkVfLsmF5
# cPuuZ2aLxAoFMvk8x7wZWGfsCqV2ciPdlk+av8nnhEwms8ozEkLvts7SB0bOV85R
# rUJ15wRzPy+/9qlzdU8I5f+T87kpNKDrmdXmTiroa2xORf9CLv9ALyJbOZVp5lUN
# 9wXNa8wnPPYpNvbDgpsxcFW4Z1vqDAZDfVHPzTbZ+UR1j9bo9gYcP5QpJppOYd1l
# Er+gnLMoAO7aDSz7FrFbyfi8ZUt5Fat6Q/MLIOEXNDoqSHFWk4Jtetaj1jO9llPh
# DXxthnxDYgaxSTkNXjND/HM4ILY25jgIxM13LyVta3dd7adpIosWc237aaoNm20G
# Tyj3UzV/9G0LB6PJCo2qxxgsRspx4wbwJcz3AUD4fsYfSTdBAzrl2rwiwkK+32ri
# rbHIL31cA4Kl9x+ClKwG1DV8l8nlAkETTJbYP0eg40Z69eAi/ydLW9Pp8yv2lE7U
# rN+APQsjYbNRYhlqdq2uBKX8QfGq2Qg=
# SIG # End signature block
