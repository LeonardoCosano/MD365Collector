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
  Purpose: Features to ease BEC investigation process on Microsoft Defender for office 365.
  Date: 2025.10.16

#>

# ToDo Remove force flag in production
Import-Module "$PSScriptRoot\Prerequisites.ps1" -Force
Import-Module "$PSScriptRoot\UnifiedAuditLog.ps1" -Force

# Title
# IsMD365CollectorReady
#
# Params
# None
#
# Description
# Checks if the environment and user executing the tool satisfy prerequisites to execute the tool
#
# Return
# Boolean. True if machine satifies all prerequisites, else false.
# 
function IsMD365CollectorReady {

    ## 1.1. Checks if powershell execution policy
    if (-not (CheckPowershellExecutionPolicy)){
        $success = SetPowershellExecutionPolicy
        if ($success = $false){
            return $false
            exit 1
        }
    }

    ## 1.2. Checks if powershell modules required by the tool are installed on system
    if (-not (CheckPowershellModulesAvailability)){
        $InstallationSuccess = InstallRequiredPowershellModules
        if ($InstallationSuccess = $false){
            return $false
            exit 1
        }
    }

    ## 1.3 Check if tool is being run on powershell ISE, environment which does not allow ExchangeOnline authentication process to be completed properly
    if (CheckPowershellIse){
        return $false
        exit 1
    }

    ## 1.4. Load required modules
    if (-not (ImportRequiredModules)){
        return $false
        exit 1
    }

    ## ToDo Authenticate as app
    ## 1.5. Authenticate
    if (-not (AuthenticateAsUserInExchangeOnline)){
        exit 1
    }

    ## ToDo Check if user account has permissions enough   
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
