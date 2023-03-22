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

    Core features included:
        - create/update the DCRs and tables automatically - based on the source object schema
        - validate the schema for naming convention issues. If exist found, it will mitigate the issues
        - update schema of DCRs and tables, if the structure of the source object changes
        - auto-fix if something goes wrong with a DCR or table
        - can remove data from the source object, if there are colums of data you don't want to send
        - can convert source objects based on CIM or PS objects into PSCustomObjects/array
        - can add relevant information to each record like UserLoggedOn, Computer, CollectionTime
    
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


$LogFile = [System.Environment]::GetEnvironmentVariable('TEMP','Machine') + "\ClientInspector.txt"
Try
    {
        Stop-Transcript   # if running
        Start-Transcript -Path $LogFile -IncludeInvocationHeader
    }
Catch
    {
    }


Write-Output ""
Write-Output "ClientInspector | Inventory of Operational & Security-related information"
Write-Output "Developed by Morten Knudsen, Microsoft MVP"
Write-Output ""
  
##########################################
# VARIABLES
##########################################

<# ----- onboarding lines ----- BEGIN #>









# On some computers doing a WMI query after applications is REALLY slow. When testing, it is nice to be able to disable this collection temporarely
$Collect_Applications_WMI                   = $false

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

    # directory where the script was started
    $ScriptDirectory = $PSScriptRoot

    switch ($Function)
        {   
            "Download"            # Typically used in Microsoft Intune environments
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

            # Used by Morten Knudsen for development
            "DevMode"
                {
                    If (Test-Path "$Env:OneDrive\Documents\GitHub\AzLogDcrIngestPS-Dev\AzLogDcrIngestPS.psm1")
                        {
                            Import-module "$Env:OneDrive\Documents\GitHub\AzLogDcrIngestPS-Dev\AzLogDcrIngestPS.psm1" -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
                        }
                    Else
                        {
                            Write-Output "Required Powershell function was NOT found .... terminating !"
                            break
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

        Try
            {
                Import-Module -Name PSWindowsUpdate
            }
        Catch
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
            $OfficeVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue

                #-------------------------------------------------------------------------------------------
                # Preparing data structure
                #-------------------------------------------------------------------------------------------
                If ($OfficeVersion)
                    {
                        $FoundO365Office = $True

                        $DataVariable = $OfficeVersion

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
        <#
                                        "4624;Microsoft-Windows-Security-Auditing"  # Succesful User Account Login
        #>

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
            
        # Convert from FileTime
        # [datetime]::FromFileTime(([Int64] $startTimeHi -shl 32) -bor $startTimeLo)

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

Try
    {
        Stop-Transcript
    }
Catch
    {
    }

