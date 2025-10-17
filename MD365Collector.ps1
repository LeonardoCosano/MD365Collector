<#

 /$$      /$$ /$$$$$$$   /$$$$$$   /$$$$$$  /$$$$$$$   /$$$$$$            /$$ /$$                       /$$                        
| $$$    /$$$| $$__  $$ /$$__  $$ /$$__  $$| $$____/  /$$__  $$          | $$| $$                      | $$                        
| $$$$  /$$$$| $$  \ $$|__/  \ $$| $$  \__/| $$      | $$  \__/  /$$$$$$ | $$| $$  /$$$$$$   /$$$$$$$ /$$$$$$    /$$$$$$   /$$$$$$ 
| $$ $$/$$ $$| $$  | $$   /$$$$$/| $$$$$$$ | $$$$$$$ | $$       /$$__  $$| $$| $$ /$$__  $$ /$$_____/|_  $$_/   /$$__  $$ /$$__  $$
| $$  $$$| $$| $$  | $$  |___  $$| $$__  $$|_____  $$| $$      | $$  \ $$| $$| $$| $$$$$$$$| $$        | $$    | $$  \ $$| $$  \__/
| $$\  $ | $$| $$  | $$ /$$  \ $$| $$  \ $$ /$$  \ $$| $$    $$| $$  | $$| $$| $$| $$_____/| $$        | $$ /$$| $$  | $$| $$      
| $$ \/  | $$| $$$$$$$/|  $$$$$$/|  $$$$$$/|  $$$$$$/|  $$$$$$/|  $$$$$$/| $$| $$|  $$$$$$$|  $$$$$$$  |  $$$$/|  $$$$$$/| $$      
|__/     |__/|_______/  \______/  \______/  \______/  \______/  \______/ |__/|__/ \_______/ \_______/   \___/   \______/ |__/      
                                                                                                                                 

  Script: MD365Collector.ps1
  Author: Leonardo Cosano
  Purpose: MD365Collector tool is intended to ease BEC investigation on Microsoft Defender for office environment. In this script file you will find generic predefined ways to use it, from initial steps to check prerequisites and authentication to generic evidence collection usage.
  Date: 2025.10.16

#>

# ToDo Remove force flag in production
Import-Module "$PSScriptRoot\Prerequisites.ps1" -Force
Import-Module "$PSScriptRoot\UnifiedAuditLog.ps1" -Force

# Title
# IsEnvironmentReadyForMD365Collector
#
# Params
# None
#
# Description
# This functions is intended to be run at the very begining of MD365Collector tool usage.
# It is though to check if user's environment is ready for executing the tool as it satisfies prerequisites (execution policy, powershell modules installed, IDE).
# It wont make any change to the user's environment.
#
# Return
# Boolean. True if machine satifies all prerequisites, else false.
# 
function IsEnvironmentReadyForMD365Collector {

    $IsEnvironmentReady = $true

    ## 1.1. Checks if powershell execution policy is permissive enough.
    if (-not (isPowershellExecutionPolicyOk)){
        $IsEnvironmentReady = $false
    }

    ## 1.2. Checks if powershell modules required by the tool are installed on system.
    if (-not (arePowershellModulesAvailable)){
        $IsEnvironmentReady = $false
    }
    
    ## 1.3 Check if tool is being run on powershell ISE.
    if (isPowershellIse -silent $false){
        $IsEnvironmentReady = $false
    }

    if (-not ($IsEnvironmentReady)){
        Write-Host 'Environment does not satisfy all requirements, please fix errors described below.' -ForegroundColor DarkRed 
    }
    else{
        Write-Host 'Environment satisfy all requirements. You can now use the tool.' -ForegroundColor DarkGreen 
    }
}

# Title
# SetEnvironmentReadyForMD365Collector
#
# Params
# None
#
# Description
# It loads required powershell modules used by the tool and authenticates against exchange online service.
#
# Return
# Boolean. True if machine satifies all prerequisites, else false.
# 
function SetEnvironmentReadyForMD365Collector {

    $IsEnvironmentReady = $true

    ## 1.1. Set proper powershell execution policy
    $wasPolicyChanged = SetPowershellExecutionPolicy
    if (-not ($wasPolicyChanged)){
        $IsEnvironmentReady = $false
    }

    ## 1.2. Checks if powershell modules required by the tool are installed on system
    $werePowershellModulesInstalled = InstallRequiredPowershellModules
    if (-not ($werePowershellModulesInstalled)){
        $IsEnvironmentReady = $false
    }

    ## 1.3 Check if tool is being run on powershell ISE, environment which does not allow ExchangeOnline authentication process to be completed properly
    $wasPowershellISEDetected = isPowershellIse -silent $true
    if ($wasPowershellISEDetected){
        $IsEnvironmentReady = $false
    }

    ## 1.4 Load required modules
    $werePowershellModulesImported = ImportRequiredModules
    if (-not ($werePowershellModulesImported)){
        $IsEnvironmentReady = $false
    }

    ## ToDo offer authentication as an app
    ## 1.5. Authenticate
    $wasAuthenticationSuccess = AuthenticateAsUserInExchangeOnline
    if (-not ($wasAuthenticationSuccess)){
        $IsEnvironmentReady = $false
    }

    ## ToDo Check if user account has permissions enough
    if (-not ($IsEnvironmentReady)){
        Write-Host 'Environment is not ready, please fix errors described below.' -ForegroundColor DarkRed 
    }
    else{
        Write-Host 'Environment is ready. You can now start using the tool.' -ForegroundColor DarkGreen 
    }
      
}

# Title
# StartCollection
#
# Params
# user. mandatory string. comma separated list of userprincipalnames of the accounts to be collected.
# start. optional datetime. beginning of the timeframe to be collected.
# end. optional datetime. beginning of the timefram to be collected.
#
# Description
# Checks if the machine executing MD365Collector satifies prerequisites to execute the tool
#
# Return
# Boolean. True if machine satifies all prerequisites, else false.
# 
function StartCollection {
    param(
        [Parameter(Mandatory=$true)]
        [string]$user,

        [Parameter(Mandatory=$false)]
        [DateTime]$start,

        [Parameter(Mandatory=$false)]
        [DateTime]$end
    )


    # 1. Collect Unified Audit Logs
    if (GetAuditLogs -user $user -start $start -end $end){
        Write-Host "Saving MD365-AuditLogs into outputs folder as '$((Get-Date).ToUniversalTime().ToString("yyyy-MM-dd")).$user.MD365-AuditLogs'" -ForegroundColor DarkCyan
    } else {
        Write-Host "Microsoft Defender Audit Logs collection failed" -ForegroundColor DarkRed
    }

    # ToDo
    # 2. Collect Cloud App Activity Logs
    # 3. Collect AADSignInLogs
    # 4. Collect AADAuditLogs 
    # 5. Collect Exchange message traces
}
