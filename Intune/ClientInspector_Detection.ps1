#------------------------------------------------------------------------------------------------
Write-Output "***********************************************************************************************"
Write-Output "CLIENT INSPECTOR | SYSTEM | DETECTION FOR MICROSOFT MEM/INTUNE"
Write-Output ""
Write-Output "Purpose: Purpose of this script is to detect if client inspector should run or it has run during the last x hours"
Write-Output ""
Write-Output "Support: Morten Knudsen - mok@2linkit.net | 40 178 179"
Write-Output "***********************************************************************************************"
#------------------------------------------------------------------------------------------------

##################################
# VARIABLES
##################################

    $RunEveryHours    = 4
    $LastRun_RegPath  = "HKLM:\SOFTWARE\2LINKIT"
    $LastRun_RegKey   = "ClientInSpector_System"


##################################
# MAIN PROGRAM
##################################
   
   
    If (-not (Get-ItemProperty -Path $LastRun_RegPath -Name $LastRun_RegKey -ErrorAction SilentlyContinue))
        {
            Write-Host "Script has never run the initial collection"
	        exit 1
        } 

    Try{        
        $Now             = (Get-date)
        [datetime]$LastRunDateTime = Get-ItemPropertyValue -Path $LastRun_RegPath -Name $LastRun_RegKey
        $NextRun         = $LastRunDateTime.AddHours($RunEveryHours)
    }

    Catch{    
        # Something went wrong
        $errMsg = $_.Exception.Message
        Write-Error $errMsg
        exit 1
    }


    If ((Get-date $Now) -le (Get-date $NextRun))
        {
            Write-Output ""
            Write-Output "  Last Run                   : $($LastRunDateTime)"
            Write-Output "  Next Run Frequency (hours) : $($RunEveryHours)"
            Write-Output "  Next Run                   : $($NextRun)"
            Write-Output "  Now                        : $($Now)"
            Write-Output ""
            Write-Output "  Action: Script should not run yet - Exit 0"

            Exit 0                        
        }
    ElseIf ((Get-date $Now) -gt (Get-date $NextRun))
        {
            Write-Output ""
            Write-Output "  Last Run                   : $($LastRunDateTime)"
            Write-Output "  Next Run Frequency (hours) : $($RunEveryHours)"
            Write-Output "  Next Run                   : $($NextRun)"
            Write-Output "  Now                        : $($Now)"
            Write-Output ""
            Write-Output "  Action: Script should run now - Exit 1"

            exit 1     
        }

