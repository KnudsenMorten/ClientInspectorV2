#Requires -RunAsAdministrator
#Requires -Version 5.1

<#
    .NAME
    ClientInspector (part 2 of 2)

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
          [ValidateSet("Download","LocalPath","PsGallery")]
          [string]$Function = "PsGallery",
      [parameter(Mandatory=$false)]
          [ValidateSet("CurrentUser","AllUsers")]
          [string]$Scope = "CurrentUser"
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

$Verbose                                    = $false

<# ----- onboarding lines ----- END  #>

$LastRun_RegPath                            = "HKLM:\SOFTWARE\ClientInspector"
$LastRun_RegKey                             = "ClientInspector_System_2"

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


# SIG # Begin signature block
# MIIRgwYJKoZIhvcNAQcCoIIRdDCCEXACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUqR5Wn9stl7vwnopBwt8RczzZ
# jDqggg3jMIIG5jCCBM6gAwIBAgIQd70OA6G3CPhUqwZyENkERzANBgkqhkiG9w0B
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
# FXe9jYfEz0ifrNteT/2wTbWwt9kwDQYJKoZIhvcNAQEBBQAEggIAxclA7Bi5GnyU
# gNl9Zeq8oDT4jNJd0cjEUwjir5FmdmS+RCuwALLoIMCCr4o2k5MnYTalX13dJys1
# dpGagX3nuzHIu0wyFagX77b5Nf2/1d1YWdq1hdjYR/oJ4AyxwmjWKTpcHvLoLKMP
# 5geGS6+opx+vFHT7wTSF9NPELGPmV8uii97+Swsks5j8dPZJwHe9lFil2dlGFiJE
# dGtqB80AZJxp7wHyXcOclCws7/++v4XB8v9yRqU4Mg77Nw/qtWCEKhwq4u064gSa
# RC4pwe1uX0MFJxGQZVOAwHqTVhXSPCX3o1Ft9uOsDGIWCKZchVOIN/VAdpC/RGZj
# C5o/mMJAE8mGtaTY3N1UFMQ46ZBEQhn6CRxJDDyVe63no6g//AfgIwETSbCCOEaf
# KDfJbR9nhyvm9x3EI61mM3f5GJYOINj//4gwEEujuir9YBAzH1bPuYOTGzuxO2OG
# H+ypIXVYgrftPUKXSlt8mMhvbOv+EvYwEX180no4cd3PeLLlGnw+BXiCnq4drPaX
# SNvHwAa+wCnYuhRMVgAAM33mCaHTLUWQFcKRF2kGnlawzNlVL7VTNPaoRoX/qtVD
# QfStth4VuU8nQ3pcsnFB2yId7n4BYULCNLdMMetrMI4ZGvVKzFJO8VO96kbESM+5
# 0/hYEhUGWofRhluY4ZDnqAO3mSn08AY=
# SIG # End signature block
