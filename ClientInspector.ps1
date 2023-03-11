﻿#Requires -RunAsAdministrator
#Requires -Version 5.0

<#
    .NAME
    ClientInspector

    .SYNOPSIS
    This script will collect lots of informations from the client. Data is sent to Azure LogAnalytics Custom Tables.
    The upload happens via Log Ingestion API, Azure Data Collection Rules (DCR) and Azure Data Collection Endpoints

    .NOTES
    VERSION: 230309
    Developed by Morten Knudsen, Microsoft MVP

    .COPYRIGHT
    @mortenknudsendk on Twitter
    Blog: https://mortenknudsen.net
    
    .LICENSE
    Licensed under the MIT license.

    .WARRANTY
    Use at your own risk, no warranty given!
#>

$LogFile = [System.Environment]::GetEnvironmentVariable('TEMP','Machine') + "\ClientInspector.txt"
Try
    {
        Stop-Transcript   # if running
        Start-Transcript -Path $LogFile -IncludeInvocationHeader
    }
Catch
    {
        Write-output "Could not start transscript logging"
    }

  
##########################################
# VARIABLES
##########################################

<# ----- onboarding lines ----- BEGIN #>

<#
    $TenantId                                   = "xxxx" 
    $LogIngestAppId                             = "xxxx" 
    $LogIngestAppSecret                         = "xxxx" 

    $DceName                                    = "xxxx" 
    $LogAnalyticsWorkspaceResourceId            = "xxxx"

    $AzDcrPrefixClient                          = "clt" 
    $AzDcrSetLogIngestApiAppPermissionsDcrLevel = $false
    $AzDcrLogIngestServicePrincipalObjectId     = "xxxx" 
    $AzDcrDceTableCreateFromReferenceMachine    = @()
    $AzDcrDceTableCreateFromAnyMachine          = $false

#>

<#  ----- onboading lines -----  END  #>

    # latest run info - needed for Intune detection to work
    $LastRun_RegPath                                 = "HKLM:\SOFTWARE\2LINKIT"
    $LastRun_RegKey                                  = "ClientInspector_System"

    # default variables
    $VerbosePreference                               = "SilentlyContinue"  # Stop, Inquire, Continue, SilentlyContinue
    $DNSName                                         = (Get-CimInstance win32_computersystem).DNSHostName +"." + (Get-CimInstance win32_computersystem).Domain
    $ComputerName                                    = (Get-CimInstance win32_computersystem).DNSHostName
    [datetime]$CollectionTime                        = ( Get-date ([datetime]::Now.ToUniversalTime()) -format "yyyy-MM-ddTHH:mm:ssK" )


############################################################################################################################################
# FUNCTIONS
############################################################################################################################################

    $ScriptDirectory = $PSScriptRoot

    # Needed funtions found in local path - for example if deployed through ConfigMgr, where we don't need to download the functions each time
    If (Test-Path "$($ScriptDirectory)\AzLogDcrIngestPS.psm1")
        {
            Import-module "$($ScriptDirectory)\AzLogDcrIngestPS.psm1" -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
        }
    
    # Used by Morten Knudsen for development
    ElseIf ("$Env:OneDrive\Documents\GitHub\AzLogDcrIngestPS-Dev\AzLogDcrIngestPS.psm1")
        {
            Import-module "$Env:OneDrive\Documents\GitHub\AzLogDcrIngestPS-Dev\AzLogDcrIngestPS.psm1" -Global -Force -DisableNameChecking
        }

    # force download using Github. This is needed for Intune remediations, since the functions library are large, and Intune only support 200 Kb at the moment
    Else   
        {
            Write-Output ""
            Write-Output "Downloading latest version of needed Powershell functions from Morten Knudsen Github .... Please Wait !"
            Write-Output ""
            Write-Output "The Powershell functions, AzLogDcrIngestPS, are developed and maintained by Morten Knudsen, Microsoft MVP"
            Write-Output ""
            Write-Output "Please send feedback or comments to mok@mortenknudsen.net. Also feel free to pull findings/issues in Github."
            Write-Output ""

            $Download = (New-Object System.Net.WebClient).DownloadFile("https://raw.githubusercontent.com/KnudsenMorten/AzLogDcrIngestPS/main/AzLogDcrIngestPS.psm1", "$($ScriptDirectory)\AzLogDcrIngestPS.psm1")  
            Import-module "$($ScriptDirectory)\AzLogDcrIngestPS.psm1" -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
        }


###############################################################
# Global Variables
#
# Used to mitigate throttling in Azure Resource Graph
# Needs to be loaded after load of functions
###############################################################

    # building global variable with all DCEs, which can be viewed by Log Ingestion app
    $global:AzDceDetails = Get-AzDceListAll -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId
    
    # building global variable with all DCRs, which can be viewed by Log Ingestion app
    $global:AzDcrDetails = Get-AzDcrListAll -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
    Write-output "Collecting User information [1]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------

        $TableName = 'InvClientComputerUserLoggedOnV2'
        $DcrName   = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------

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
        $DataVariable1 = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

        # add Computer info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

    #-------------------------------------------------------------------------------------------
    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
    #-------------------------------------------------------------------------------------------

        CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                             -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                             -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                             -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                             -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                             -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                             -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine


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
            $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting Bios information"

            $DataVariable = Get-CimInstance -ClassName Win32_BIOS

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------

            # convert CIM array to PSCustomObject and remove CIM class information
            $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable
    
            # add CollectionTime to existing array
            $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

            # add Computer & UserLoggedOn info to existing array
            $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

            # Validating/fixing schema data structure of source data
            $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

            # Aligning data structure with schema (requirement for DCR)
            $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

    #-------------------------------------------------------------------------------------------
    # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
    #-------------------------------------------------------------------------------------------

        CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                             -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                             -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                             -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                             -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                             -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                             -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
    #-----------------------------------------------------------------------------------------------
    # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
    #-----------------------------------------------------------------------------------------------

        Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                           -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


    ####################################
    # Processor
    ####################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientComputerInfoProcessorV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting Processor information"
            $DataVariable = Get-CimInstance -ClassName Win32_Processor | Select-Object -ExcludeProperty "CIM*"

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------

            # convert CIM array to PSCustomObject and remove CIM class information
            $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable
    
            # add CollectionTime to existing array
            $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

            # add Computer & UserLoggedOn info to existing array
            $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

            # Validating/fixing schema data structure of source data
            $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

            # Aligning data structure with schema (requirement for DCR)
            $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


    ####################################
    # Computer System
    ####################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientComputerInfoSystemV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting Computer system information"

            $DataVariable = Get-CimInstance -ClassName Win32_ComputerSystem

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------

            # convert CIM array to PSCustomObject and remove CIM class information
            $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable
    
            # add CollectionTime to existing array
            $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

            # add Computer & UserLoggedOn info to existing array
            $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

            # Validating/fixing schema data structure of source data
            $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

            # Aligning data structure with schema (requirement for DCR)
            $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId

    ####################################
    # Computer Info
    ####################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientComputerInfoV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting computer information"

            $DataVariable = Get-ComputerInfo

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------

            # convert CIM array to PSCustomObject and remove CIM class information
            $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable
    
            # add CollectionTime to existing array
            $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

            # add Computer & UserLoggedOn info to existing array
            $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

            # Validating/fixing schema data structure of source data
            $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

            # Aligning data structure with schema (requirement for DCR)
            $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


    ####################################
    # OS Info
    ####################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientComputerOSInfoV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting OS information"

            $DataVariable = Get-CimInstance -ClassName Win32_OperatingSystem

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------

            # convert CIM array to PSCustomObject and remove CIM class information
            $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable
    
            # add CollectionTime to existing array
            $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

            # add Computer & UserLoggedOn info to existing array
            $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

            # Validating/fixing schema data structure of source data
            $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

            # Aligning data structure with schema (requirement for DCR)
            $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


    ####################################
    # Last Restart
    ####################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientComputerInfoLastRestartV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output ""
            Write-Output "Collecting Last restart information"

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
            $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

            # add Computer & UserLoggedOn info to existing array
            $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

            # Validating/fixing schema data structure of source data
            $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

            # Aligning data structure with schema (requirement for DCR)
            $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


###############################################################
# APPLICATIONS (WMI) [3]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    write-output "INSTALLED APPLICATIONS INFORMATION [3]"
    Write-output ""
    write-output "WMI information about installed applications"
    Write-output ""

    #------------------------------------------------
    # Installed Application (WMI)
    #------------------------------------------------

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientApplicationsFromWmiV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output "Collecting installed application information via WMI (slow)"

            $DataVariable = Get-CimInstance -Class Win32_Product

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------
    
        # convert Cim object and remove PS class information
        $DataVariable = Convert-CimArrayToObjectFixStructure -Data $DataVariable

        # add CollectionTime to existing array
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

        # add Computer & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

        # Get insight about the schema structure of an object BEFORE changes. Command is only needed to verify columns in schema
        # $SchemaBefore = Get-ObjectSchemaAsArray -Data $DataVariable
        
        # Remove unnecessary columns in schema
        $DataVariable = Filter-ObjectExcludeProperty -Data $DataVariable -ExcludeProperty __*,SystemProperties,Scope,Qualifiers,Properties,ClassPath,Class,Derivation,Dynasty,Genus,Namespace,Path,Property_Count,RelPath,Server,Superclass

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


###############################################################
# APPLICATIONS (REGISTRY) [3]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    write-output "INSTALLED APPLICATIONS INFORMATION [3]"
    Write-output ""
    write-output "Registry information about installed applications"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientApplicationsFromRegistryV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output "Collecting installed applications information via registry"

        $UninstallValuesX86 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue
        $UninstallValuesX64 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue

        $DataVariable       = $UninstallValuesX86
        $DataVariable      += $UninstallValuesX64

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        # removing apps without DisplayName fx KBs
        $DataVariable = $DataVariable | Where-Object { $_.DisplayName -ne $null }

        # convert PS object and remove PS class information
        $DataVariable = Convert-PSArrayToObjectFixStructure -Data $DataVariable

        # add CollectionTime to existing array
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

        # add Computer & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

        # Get insight about the schema structure of an object BEFORE changes. Command is only needed to verify columns in schema
        # $SchemaBefore = Get-ObjectSchemaAsArray -Data $DataVariable
        
        # Remove unnecessary columns in schema
        $DataVariable = Filter-ObjectExcludeProperty -Data $DataVariable -ExcludeProperty Memento*,Inno*,'(default)',1033

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output "Collecting antivirus information"

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
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

        # add Computer & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output "Collecting Microsoft Defender Antivirus information"

        $MPComputerStatus = Get-MpComputerStatus
        $MPPreference = Get-MpPreference

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        If ($MPComputerStatus) 
            {
                $MPComputerStatusObject = [PSCustomObject]@{
                                                                MPComputerStatusFound = $True
                                                            }
            }
        Else
            {
                $MPComputerStatusObject = [PSCustomObject]@{
                                                                MPComputerStatusFound = $false
                                                            }
            }

    # Collecting Defender AV MPPreference-settings
        $MPPreference = Get-MpPreference
        If ($MPPreference) 
            {
                $MPPreferenceObject = [PSCustomObject]@{
                                                            MPPreferenceFound = $True
                                                        }
            }
        Else
            {
                $MPPreferenceObject = [PSCustomObject]@{
                                                            MPPreferenceFound = $False
                                                        }
            }

    # Preparing data
        $DataVariable = [PSCustomObject]@{
            MPComputerStatusFound                         = $MPComputerStatusObject.MPComputerStatusFound
            MPPreferenceFound                             = $MPPreferenceObject.MPPreferenceFound
            AMEngineVersion                               = $MPComputerStatus.AMEngineVersion
            AMProductVersion                              = $MPComputerStatus.AMProductVersion
            AMRunningMode                                 = $MPComputerStatus.AMRunningMode
            AMServiceEnabled                              = $MPComputerStatus.AMServiceEnabled
            AMServiceVersion                              = $MPComputerStatus.AMServiceVersion
            AntispywareEnabled                            = $MPComputerStatus.AntispywareEnabled
            AntispywareSignatureAge                       = $MPComputerStatus.AntispywareSignatureAge
            AntispywareSignatureLastUpdated               = $MPComputerStatus.AntispywareSignatureLastUpdated
            AntispywareSignatureVersion                   = $MPComputerStatus.AntispywareSignatureVersion
            AntivirusEnabled                              = $MPComputerStatus.AntivirusEnabled
            AntivirusSignatureAge                         = $MPComputerStatus.AntivirusSignatureAge
            AntivirusSignatureLastUpdated                 = $MPComputerStatus.AntivirusSignatureLastUpdated
            AntivirusSignatureVersion                     = $MPComputerStatus.AntivirusSignatureVersion
            BehaviorMonitorEnabled                        = $MPComputerStatus.BehaviorMonitorEnabled
            DefenderSignaturesOutOfDate                   = $MPComputerStatus.DefenderSignaturesOutOfDate
            DeviceControlDefaultEnforcement               = $MPComputerStatus.DeviceControlDefaultEnforcement
            DeviceControlPoliciesLastUpdated              = $MPComputerStatus.DeviceControlPoliciesLastUpdated
            DeviceControlState                            = $MPComputerStatus.DeviceControlState
            FullScanAge                                   = $MPComputerStatus.FullScanAge
            FullScanEndTime                               = $MPComputerStatus.FullScanEndTime
            FullScanOverdue                               = $MPComputerStatus.FullScanOverdue
            FullScanRequired                              = $MPComputerStatus.FullScanRequired
            FullScanSignatureVersion                      = $MPComputerStatus.FullScanSignatureVersion
            FullScanStartTime                             = $MPComputerStatus.FullScanStartTime
            IoavProtectionEnabled                         = $MPComputerStatus.IoavProtectionEnabled
            IsTamperProtected                             = $MPComputerStatus.IsTamperProtected
            IsVirtualMachine                              = $MPComputerStatus.IsVirtualMachine
            LastFullScanSource                            = $MPComputerStatus.LastFullScanSource
            LastQuickScanSource                           = $MPComputerStatus.LastQuickScanSource
            NISEnabled                                    = $MPComputerStatus.NISEnabled
            NISEngineVersion                              = $MPComputerStatus.NISEngineVersion
            NISSignatureAge                               = $MPComputerStatus.NISSignatureAge
            NISSignatureLastUpdated                       = $MPComputerStatus.NISSignatureLastUpdated
            NISSignatureVersion                           = $MPComputerStatus.NISSignatureVersion
            OnAccessProtectionEnabled                     = $MPComputerStatus.OnAccessProtectionEnabled
            ProductStatus                                 = $MPComputerStatus.ProductStatus
            QuickScanAge                                  = $MPComputerStatus.QuickScanAge
            QuickScanEndTime                              = $MPComputerStatus.QuickScanEndTime
            QuickScanOverdue                              = $MPComputerStatus.QuickScanOverdue
            QuickScanSignatureVersion                     = $MPComputerStatus.QuickScanSignatureVersion
            QuickScanStartTime                            = $MPComputerStatus.QuickScanStartTime
            RealTimeProtectionEnabled                     = $MPComputerStatus.RealTimeProtectionEnabled
            RealTimeScanDirection                         = $MPComputerStatus.RealTimeScanDirection
            RebootRequired                                = $MPComputerStatus.RebootRequired
            TamperProtectionSource                        = $MPComputerStatus.TamperProtectionSource
            TDTMode                                       = $MPComputerStatus.TDTMode
            TDTStatus                                     = $MPComputerStatus.TDTStatus
            TDTTelemetry                                  = $MPComputerStatus.TDTTelemetry
            TroubleShootingDailyMaxQuota                  = $MPComputerStatus.TroubleShootingDailyMaxQuota
            TroubleShootingDailyQuotaLeft                 = $MPComputerStatus.TroubleShootingDailyQuotaLeft
            TroubleShootingEndTime                        = $MPComputerStatus.TroubleShootingEndTime
            TroubleShootingExpirationLeft                 = $MPComputerStatus.TroubleShootingExpirationLeft
            TroubleShootingMode                           = $MPComputerStatus.TroubleShootingMode
            TroubleShootingModeSource                     = $MPComputerStatus.TroubleShootingModeSource
            TroubleShootingQuotaResetTime                 = $MPComputerStatus.TroubleShootingQuotaResetTime
            TroubleShootingStartTime                      = $MPComputerStatus.TroubleShootingStartTime
            AllowDatagramProcessingOnWinServer            = $MPPreference.AllowDatagramProcessingOnWinServer
            AllowNetworkProtectionDownLevel               = $MPPreference.AllowNetworkProtectionDownLevel
            AllowNetworkProtectionOnWinServer             = $MPPreference.AllowNetworkProtectionOnWinServer
            AllowSwitchToAsyncInspection                  = $MPPreference.AllowSwitchToAsyncInspection
            AttackSurfaceReductionOnlyExclusions          = $MPPreference.AttackSurfaceReductionOnlyExclusions
            AttackSurfaceReductionRules_Actions           = $MPPreference.AttackSurfaceReductionRules_Actions
            AttackSurfaceReductionRules_Ids               = $MPPreference.AttackSurfaceReductionRules_Ids
            CheckForSignaturesBeforeRunningScan           = $MPPreference.CheckForSignaturesBeforeRunningScan 
            CloudBlockLevel                               = $MPPreference.CloudBlockLevel 
            CloudExtendedTimeout                          = $MPPreference.CloudExtendedTimeout
            ComputerID                                    = $MPPreference.ComputerID
            ControlledFolderAccessAllowedApplications     = $MPPreference.ControlledFolderAccessAllowedApplications
            ControlledFolderAccessProtectedFolders        = $MPPreference.ControlledFolderAccessProtectedFolders
            DefinitionUpdatesChannel                      = $MPPreference.DefinitionUpdatesChannel
            DisableArchiveScanning                        = $MPPreference.DisableArchiveScanning
            DisableAutoExclusions                         = $MPPreference.DisableAutoExclusions
            DisableBehaviorMonitoring                     = $MPPreference.DisableBehaviorMonitoring
            DisableBlockAtFirstSeen                       = $MPPreference.DisableBlockAtFirstSeen
            DisableCatchupFullScan                        = $MPPreference.DisableCatchupFullScan
            DisableCatchupQuickScan                       = $MPPreference.DisableCatchupQuickScan
            DisableCpuThrottleOnIdleScans                 = $MPPreference.DisableCpuThrottleOnIdleScans
            DisableDatagramProcessing                     = $MPPreference.DisableDatagramProcessing 
            DisableDnsOverTcpParsing                      = $MPPreference.DisableDnsOverTcpParsing
            DisableDnsParsing                             = $MPPreference.DisableDnsParsing
            DisableEmailScanning                          = $MPPreference.DisableEmailScanning
            DisableFtpParsing                             = $MPPreference.DisableFtpParsing
            DisableGradualRelease                         = $MPPreference.DisableGradualRelease 
            DisableHttpParsing                            = $MPPreference.DisableHttpParsing
            DisableInboundConnectionFiltering             = $MPPreference.DisableInboundConnectionFiltering 
            DisableIOAVProtection                         = $MPPreference.DisableIOAVProtection
            DisableNetworkProtectionPerfTelemetry         = $MPPreference.DisableNetworkProtectionPerfTelemetry
            DisablePrivacyMode                            = $MPPreference.DisablePrivacyMode
            DisableRdpParsing                             = $MPPreference.DisableRdpParsing
            DisableRealtimeMonitoring                     = $MPPreference.DisableRealtimeMonitoring
            DisableRemovableDriveScanning                 = $MPPreference.DisableRemovableDriveScanning
            DisableRestorePoint                           = $MPPreference.DisableRestorePoint
            DisableScanningMappedNetworkDrivesForFullScan = $MPPreference.DisableScanningMappedNetworkDrivesForFullScan
            DisableScanningNetworkFiles                   = $MPPreference.DisableScanningNetworkFiles
            DisableScriptScanning                         = $MPPreference.DisableScriptScanning
            DisableSshParsing                             = $MPPreference.DisableSshParsing
            DisableTDTFeature                             = $MPPreference.DisableTDTFeature
            DisableTlsParsing                             = $MPPreference.DisableTlsParsing
            EnableControlledFolderAccess                  = $MPPreference.EnableControlledFolderAccess
            EnableDnsSinkhole                             = $MPPreference.EnableDnsSinkhole
            EnableFileHashComputation                     = $MPPreference.EnableFileHashComputation 
            EnableFullScanOnBatteryPower                  = $MPPreference.EnableFullScanOnBatteryPower
            EnableLowCpuPriority                          = $MPPreference.EnableLowCpuPriority
            EnableNetworkProtection                       = $MPPreference.EnableNetworkProtection
            EngineUpdatesChannel                          = $MPPreference.EngineUpdatesChannel
            ExclusionExtension                            = $MPPreference.ExclusionExtension
            ExclusionIpAddress                            = $MPPreference.ExclusionIpAddress
            ExclusionPath                                 = $MPPreference.ExclusionPath
            ExclusionProcess                              = $MPPreference.ExclusionProcess
            ForceUseProxyOnly                             = $MPPreference.ForceUseProxyOnly
            HighThreatDefaultAction                       = $MPPreference.HighThreatDefaultAction
            LowThreatDefaultAction                        = $MPPreference.LowThreatDefaultAction
            MAPSReporting                                 = $MPPreference.MAPSReporting
            MeteredConnectionUpdates                      = $MPPreference.MeteredConnectionUpdates
            ModerateThreatDefaultAction                   = $MPPreference.ModerateThreatDefaultAction
            PlatformUpdatesChannel                        = $MPPreference.PlatformUpdatesChannel 
            ProxyBypass                                   = $MPPreference.ProxyBypass
            ProxyPacUrl                                   = $MPPreference.ProxyPacUrl
            ProxyServer                                   = $MPPreference.ProxyServer
            PUAProtection                                 = $MPPreference.PUAProtection 
            QuarantinePurgeItemsAfterDelay                = $MPPreference.QuarantinePurgeItemsAfterDelay
            RandomizeScheduleTaskTimes                    = $MPPreference.RandomizeScheduleTaskTimes
            RemediationScheduleDay                        = $MPPreference.RemediationScheduleDay
            RemediationScheduleTime                       = $MPPreference.RemediationScheduleTime
            ReportingAdditionalActionTimeOut              = $MPPreference.ReportingAdditionalActionTimeOut
            ReportingCriticalFailureTimeOut               = $MPPreference.ReportingCriticalFailureTimeOut
            ReportingNonCriticalTimeOut                   = $MPPreference.ReportingNonCriticalTimeOut
            ScanAvgCPULoadFactor                          = $MPPreference.ScanAvgCPULoadFactor 
            ScanOnlyIfIdleEnabled                         = $MPPreference.ScanOnlyIfIdleEnabled
            ScanParameters                                = $MPPreference.ScanParameters
            ScanPurgeItemsAfterDelay                      = $MPPreference.ScanPurgeItemsAfterDelay
            ScanScheduleDay                               = $MPPreference.ScanScheduleDay
            ScanScheduleOffset                            = $MPPreference.ScanScheduleOffset
            ScanScheduleQuickScanTime                     = $MPPreference.ScanScheduleQuickScanTime
            ScanScheduleTime                              = $MPPreference.ScanScheduleTime 
            SchedulerRandomizationTime                    = $MPPreference.SchedulerRandomizationTime
            ServiceHealthReportInterval                   = $MPPreference.ServiceHealthReportInterval
            SevereThreatDefaultAction                     = $MPPreference.SevereThreatDefaultAction
            SharedSignaturesPath                          = $MPPreference.SharedSignaturesPath
            SignatureAuGracePeriod                        = $MPPreference.SignatureAuGracePeriod
            SignatureBlobFileSharesSources                = $MPPreference.SignatureBlobFileSharesSources
            SignatureBlobUpdateInterval                   = $MPPreference.SignatureBlobUpdateInterval
            SignatureDefinitionUpdateFileSharesSources    = $MPPreference.SignatureDefinitionUpdateFileSharesSources
            SignatureDisableUpdateOnStartupWithoutEngine  = $MPPreference.SignatureDisableUpdateOnStartupWithoutEngine
            SignatureFallbackOrder                        = $MPPreference.SignatureFallbackOrder
            SignatureFirstAuGracePeriod                   = $MPPreference.SignatureFirstAuGracePeriod
            SignatureScheduleDay                          = $MPPreference.SignatureScheduleDay
            SignatureScheduleTime                         = $MPPreference.SignatureScheduleTime
            SignatureUpdateCatchupInterval                = $MPPreference.SignatureUpdateCatchupInterval
            SignatureUpdateInterval                       = $MPPreference.SignatureUpdateInterval
            SubmitSamplesConsent                          = $MPPreference.SubmitSamplesConsent
            ThreatIDDefaultAction_Actions                 = $MPPreference.ThreatIDDefaultAction_Actions
            ThreatIDDefaultAction_Ids                     = $MPPreference.ThreatIDDefaultAction_Ids
            ThrottleForScheduledScanOnly                  = $MPPreference.ThrottleForScheduledScanOnly
            TrustLabelProtectionStatus                    = $MPPreference.TrustLabelProtectionStatus
            UILockdown                                    = $MPPreference.UILockdown 
            UnknownThreatDefaultAction                    = $MPPreference.UnknownThreatDefaultAction
        }
    
        # add CollectionTime to existing array
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

        # add Computer & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"


    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
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
                        $DataVariable = Convert-PSArrayToObjectFixStructure -data $DataVariable

                        # add CollectionTime to existing array
                        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

                        # add Computer & UserLoggedOn info to existing array
                        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                        # Validating/fixing schema data structure of source data
                        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                        # Aligning data structure with schema (requirement for DCR)
                        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
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
                        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

                        # add Computer & UserLoggedOn info to existing array
                        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                        # Validating/fixing schema data structure of source data
                        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                        # Aligning data structure with schema (requirement for DCR)
                        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
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
                        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

                        # add Computer & UserLoggedOn info to existing array
                        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                        # Validating/fixing schema data structure of source data
                        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                        # Aligning data structure with schema (requirement for DCR)
                        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
                    }

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output "Collecting VPN information"

        # Default Values
            $VPNSoftware = ""
            $VPNVersion = ""

        # Checking
            ForEach ($Application in $InstalledApplications)
                {
                    #-----------------------------------------
                    # Looking for Cisco AnyConnect
                    #-----------------------------------------
                    If ( ($Application.Vendor -like 'Cisco*') -and ($Application.name -like "*AnyConnect*") )
                        {
                            $VPNSoftware = $Application.Name
                            $VPNVersion = $Application.Version
                        }

                    #-----------------------------------------
                    # Looking for Palo Alto
                    #-----------------------------------------
                    If ( ($Application.Vendor -like 'Palo Alto*') -and ($Application.name -like "*Global*") )
                        {
                            $VPNSoftware = $Application.Name
                            $VPNVersion = $Application.Version
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
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

        # add Computer & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output "Collecting LAPS information"

        # Default Values
            $LAPSSoftware = ""
            $LAPSVersion = ""

        # Checking
            ForEach ($Application in $InstalledApplications)
                {
                    #-----------------------------------------
                    # Looking for LAPS
                    #-----------------------------------------
                    If ( ($Application.Vendor -like 'Microsoft*') -and ($Application.name -like "*Local Administrator Password*") )
                        {
                            $LAPSSoftware = $Application.Name
                            $LAPSVersion = $Application.Version
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
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

        # add Computer & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output "Collecting Admin By Request information"

        # Default Values
        $ABRSoftware = ""
        $ABRVersion = ""

        ForEach ($Application in $InstalledApplications)
            {
                #-----------------------------------------
                # Looking for Admin By Request
                #-----------------------------------------
                If ( ($Application.Vendor -like 'FastTrack*') -and ($Application.name -like "*Admin By Request*") )
                    {
                        $ABRSoftware = $Application.Name
                        $ABRVersion = $Application.Version
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
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
            }

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
            $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output "Collecting Windows Update Last Results information"

            $DataVariable = Get-WULastResults

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------

            # convert CIM array to PSCustomObject and remove CIM class information
            $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable
    
            # add CollectionTime to existing array
            $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

            # add Computer & UserLoggedOn info to existing array
            $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

            # Validating/fixing schema data structure of source data
            $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

            # Aligning data structure with schema (requirement for DCR)
            $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


    #################################################
    # Windows Update Source Information
    #################################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientWindowsUpdateServiceManagerV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output "Collecting Windows Update Source Information"

            $DataVariable = Get-WUServiceManager | Where-Object { $_.IsDefaultAUService -eq $true }

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------

            # convert CIM array to PSCustomObject and remove CIM class information
            $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable
    
            # add CollectionTime to existing array
            $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

            # add Computer & UserLoggedOn info to existing array
            $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

            # Validating/fixing schema data structure of source data
            $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

            # Aligning data structure with schema (requirement for DCR)
            $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


    #################################################
    # Pending Windows Updates
    #################################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientWindowsUpdatePendingUpdatesV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output "Collecting Pending Windows Updates Information"

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
                    # convert CIM array to PSCustomObject and remove CIM class information
                    $DataVariable = Convert-CimArrayToObjectFixStructure -data $WU_PendingUpdates

                    # add CollectionTime to existing array
                    $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $WU_PendingUpdates

                    # add Computer & UserLoggedOn info to existing array
                    $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                    # Remove unnecessary columns in schema
                    $DataVariable = Filter-ObjectExcludeProperty -Data $DataVariable -ExcludeProperty BundledUpdates, Categories, CveIds, DownloadContents, Identity, InstallationBehavior, UninstallationSteps, `
                                                                                                      KBArticleIDs, Languages, MoreInfoUrls, SecurityBulletinIDs, SuperSededUpdateIds, UninstallationBehavior, WindowsDriverUpdateEntries

                    # Validating/fixing schema data structure of source data
                    $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                    # Aligning data structure with schema (requirement for DCR)
                    $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
                }

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


    #############################################################
    # Status of Windows Update installations during last 31 days
    #############################################################

        #-------------------------------------------------------------------------------------------
        # Variables
        #-------------------------------------------------------------------------------------------
            
            $TableName  = 'InvClientWindowsUpdateLastInstallationsV2'   # must not contain _CL
            $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output "Collecting Last Installations of Windows Updates information"

            $WU_LastInstallations = Get-WuHistory -MaxDate ((Get-date).AddDays(-31))

        #-------------------------------------------------------------------------------------------
        # Preparing data structure
        #-------------------------------------------------------------------------------------------

            # Add CollectionTime & ComputerName to array
            If ($WU_LastInstallations)
                {
                    $CountVar = ($WU_LastInstallations | Measure-Object).count
                    
                    $PosDataVariable = 0
                    Do
                        {
                            # Remove UninstallationSteps from array
                            $WU_LastInstallations[$PosDataVariable].PSObject.Properties.Remove("UninstallationSteps")

                            # Remove Categories from array
                            $WU_LastInstallations[$PosDataVariable].PSObject.Properties.Remove("Categories")
    
                            # Remove UpdateIdentity from array
                            $WU_LastInstallations[$PosDataVariable].PSObject.Properties.Remove("UpdateIdentity")

                            $PosDataVariable += 1
                        }
                    Until ($PosDataVariable -eq $CountVar)

                    # convert CIM array to PSCustomObject and remove CIM class information
                    $DataVariable = Convert-CimArrayToObjectFixStructure -data $WU_LastInstallations

                    # add CollectionTime to existing array
                    $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

                    # add Computer & UserLoggedOn info to existing array
                    $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                    # Validating/fixing schema data structure of source data
                    $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                    # Aligning data structure with schema (requirement for DCR)
                    $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
            }
        Else
            {
                $DataVariable = $WU_LastInstallations
            }

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output "Collecting Bitlocker information"

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
        $DataVariable = Convert-CimArrayToObjectFixStructure -data $DataVariable
    
        # add CollectionTime to existing array
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

        # add Computer & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output "Collecting Eventlog information"

        $FilteredEvents      = @()
        $Appl_Events_ALL     = @()
        $System_Events_ALL   = @()
        $Security_Events_ALL = @()

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

        # convert CIM array to PSCustomObject and remove CIM class information
        $DataVariable = Convert-CimArrayToObjectFixStructure -data $FilteredEvents
    
        # add CollectionTime to existing array
        $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

        # add Computer & UserLoggedOn info to existing array
        $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

        # Validating/fixing schema data structure of source data
        $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

        # Aligning data structure with schema (requirement for DCR)
        $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output "Collecting Network Adapter information"

        $NetworkAdapter = Get-NetAdapter

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------
        If ($NetworkAdapter)
            {
                # convert CIM array to PSCustomObject and remove CIM class information
                $DataVariable = Convert-CimArrayToObjectFixStructure -data $NetworkAdapter
    
                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                # Get insight about the schema structure of an object BEFORE changes. Command is only needed to verify schema - can be disabled
                $SchemaBefore = Get-ObjectSchemaAsArray -Data $DataVariable
        
                # Remove unnecessary columns in schema
                $DataVariable = Filter-ObjectExcludeProperty -Data $DataVariable -ExcludeProperty Memento*,Inno*,'(default)',1033

                # Get insight about the schema structure of an object AFTER changes. Command is only needed to verify schema - can be disabled
                $Schema = Get-ObjectSchemaAsArray -Data $DataVariable

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable            }
        Else
            {

                # log issue - typically WMI issue
                $TableName  = 'InvClientCollectionIssuesV2'   # must not contain _CL
                $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

                $DataVariable = [pscustomobject]@{
                                                   IssueCategory   = "NetworkAdapterInformation"
                                                 }

                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
            }         

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output "Collecting IPv4 information"

        $IPv4Status = Get-NetIPAddress -AddressFamily IPv4

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------
        If ($IPv4Status)
            {
                # convert CIM array to PSCustomObject and remove CIM class information
                $DataVariable = Convert-CimArrayToObjectFixStructure -data $IPv4Status
    
                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
            }
        Else
            {

                # log issue - typically WMI issue
                $TableName  = 'InvClientCollectionIssuesV2'   # must not contain _CL
                $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

                $DataVariable = [pscustomobject]@{
                                                   IssueCategory   = "IPInformation"
                                                 }

                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
            }         

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output "Collecting Local Admin information"

        $LocalAdminGroupname = (Get-localgroup -Sid S-1-5-32-544).name       # SID S-1-5-32-544 = local computer Administrators group
        $LocalAdmins = Get-LocalGroupMember -Group  $LocalAdminGroupname

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        If ($LocalAdmins -eq $null)
            {
                ########################################################################################################################
                # Fix local admin group - The problem is empty SIDs in the Administrators Group caused by domain joins/leave/join
                ########################################################################################################################
                    $administrators = @(
                    ([ADSI]"WinNT://./$($LocalAdminGroupname)").psbase.Invoke('Members') |
                    % { 
                        $_.GetType().InvokeMember('AdsPath','GetProperty',$null,$($_),$null) 
                    }
                    ) -match '^WinNT';

                    $administrators = $administrators -replace "WinNT://",""

                    foreach ($administrator in $administrators)
                        {
                            #write-host $administrator "got here"
                            if ($administrator -like "$env:COMPUTERNAME/*" -or $administrator -like "AzureAd/*")
                                {
                                    continue;
                                }
                            elseif ($administrator -match "S-1") #checking for empty/orphaned SIDs only
                                {
                                    Remove-LocalGroupMember -group $LocalAdminGroupname -member $administrator
                                }
                        }
            }
        Else
            {
                # convert PS array to PSCustomObject and remove PS class information
                $DataVariable = Convert-PSArrayToObjectFixStructure -data $LocalAdmins
    
                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
            }

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
            $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

        #-------------------------------------------------------------------------------------------
        # Collecting data (in)
        #-------------------------------------------------------------------------------------------
            
            Write-Output "Collecting Windows Firewall information"

            $WinFw = Get-NetFirewallProfile -policystore activestore

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------
        If ($WinFw)
            {
                # convert CIM array to PSCustomObject and remove CIM class information
                $DataVariable = Convert-CimArrayToObjectFixStructure -data $WinFw
    
                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
            }
        Else
            {

                # log issue - typically WMI issue
                $TableName  = 'InvClientCollectionIssuesV2'   # must not contain _CL
                $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

                $DataVariable = [pscustomobject]@{
                                                   IssueCategory   = "WinFwInformation"
                                                 }

                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
            }         

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


###############################################################
# GROUP POLICY REFRESH [17]
###############################################################

    Write-output ""
    Write-output "#########################################################################################"
    Write-output "GROUP POLICY INFORMATION [17]"
    Write-output ""

    #-------------------------------------------------------------------------------------------
    # Variables
    #-------------------------------------------------------------------------------------------
            
        $TableName  = 'InvClientGroupPolicyRefreshV2'   # must not contain _CL
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output "Collecting Group Policy information"

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
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataArray

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
            }

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
        $DcrName    = "dcr-" + $AzDcrPrefixClient + "-" + $TableName + "_CL"

    #-------------------------------------------------------------------------------------------
    # Collecting data (in)
    #-------------------------------------------------------------------------------------------
            
        Write-Output "Collecting TPM information"

        $TPM = Get-TPM -ErrorAction SilentlyContinue -WarningVariable SilentlyContinue

    #-------------------------------------------------------------------------------------------
    # Preparing data structure
    #-------------------------------------------------------------------------------------------

        If ($TPM)
            {
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
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $TPM

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable
            }
        Else
            {
                $DataVariable = [pscustomobject]@{
                                                    IssueCategory   = "TPM"
                                                    }

                # add CollectionTime to existing array
                $DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable

                # add Computer & UserLoggedOn info to existing array
                $DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName -Column2Name UserLoggedOn -Column2Data $UserLoggedOn

                # Validating/fixing schema data structure of source data
                $DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable

                # Aligning data structure with schema (requirement for DCR)
                $DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable
            }

        #-------------------------------------------------------------------------------------------
        # Create/Update Schema for LogAnalytics Table & Data Collection Rule schema
        #-------------------------------------------------------------------------------------------

            CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId  `
                                                 -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId `
                                                 -DceName $DceName -DcrName $DcrName -TableName $TableName -Data $DataVariable `
                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
        
        #-----------------------------------------------------------------------------------------------
        # Upload data to LogAnalytics using DCR / DCE / Log Ingestion API
        #-----------------------------------------------------------------------------------------------

            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $DceName -DcrName $DcrName -Data $DataVariable `
                                                               -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantId


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
