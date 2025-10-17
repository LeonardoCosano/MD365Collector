<#

 /$$      /$$ /$$$$$$$   /$$$$$$   /$$$$$$  /$$$$$$$   /$$$$$$            /$$ /$$                       /$$                        
| $$$    /$$$| $$__  $$ /$$__  $$ /$$__  $$| $$____/  /$$__  $$          | $$| $$                      | $$                        
| $$$$  /$$$$| $$  \ $$|__/  \ $$| $$  \__/| $$      | $$  \__/  /$$$$$$ | $$| $$  /$$$$$$   /$$$$$$$ /$$$$$$    /$$$$$$   /$$$$$$ 
| $$ $$/$$ $$| $$  | $$   /$$$$$/| $$$$$$$ | $$$$$$$ | $$       /$$__  $$| $$| $$ /$$__  $$ /$$_____/|_  $$_/   /$$__  $$ /$$__  $$
| $$  $$$| $$| $$  | $$  |___  $$| $$__  $$|_____  $$| $$      | $$  \ $$| $$| $$| $$$$$$$$| $$        | $$    | $$  \ $$| $$  \__/
| $$\  $ | $$| $$  | $$ /$$  \ $$| $$  \ $$ /$$  \ $$| $$    $$| $$  | $$| $$| $$| $$_____/| $$        | $$ /$$| $$  | $$| $$      
| $$ \/  | $$| $$$$$$$/|  $$$$$$/|  $$$$$$/|  $$$$$$/|  $$$$$$/|  $$$$$$/| $$| $$|  $$$$$$$|  $$$$$$$  |  $$$$/|  $$$$$$/| $$      
|__/     |__/|_______/  \______/  \______/  \______/  \______/  \______/ |__/|__/ \_______/ \_______/   \___/   \______/ |__/      
                                                                                                                                 

  Script: Prerequisites.ps1
  Author: Leonardo Cosano
  Purpose: Features & functions to check if your machine is ready to run the tool.
  Date: 2025.10.16

#>

# Title 
# Global variable powershellModulesRequiredByTool
# 
# Description
# Variable containing powershell module names which are required by MD365Collector to work.
# 
$script:powershellModulesRequiredByTool = @(
    'Microsoft.PowerShell.Utility',
    'ExchangeOnlineManagement'
    #'Az'
)

# Title
# isPowershellExecutionPolicyOk
#
# Params
# None
#
# Description
# Checks if the powershell execution policy is permissive enough for tool execution (unrestricted or bypass) and suggests fixes if it is not.
#
# Return
# Boolean. True if powershell execution policy is unrestricted or bypass. Else, false.
# 
function isPowershellExecutionPolicyOk {

    Write-Host "Checking if powershell execution policy is set to, at least, unrestricted." -ForegroundColor DarkCyan

    $policy = Get-ExecutionPolicy -Scope CurrentUser
    if ($policy -notin @('Unrestricted','Bypass')){
        Write-Host "You should set 'Unrestricted' powershell execution policy or permissive, not $policy" -ForegroundColor DarkYellow
        return $false
    }

    return $true
}


# Title
# arePowershellModulesAvailable
#
# Params
# None
#
# Description
# Checks if the machine executing MD365Collector has required powershell modules installed and suggests fixes if it is not.
#
# Return
# Boolean. True if machine has all required powershell modules installed, else false.
# 
function arePowershellModulesAvailable {

    Write-Host "Checking if all required powershell modules are installed on the system..." -ForegroundColor DarkCyan
    $allModulesInstalled = $true

    # Iterate over list of required modules
    foreach ($module in $powershellModulesRequiredByTool) {

        #check if module is installed
        $isModuleInstalled = Get-Module -ListAvailable -Name $module

        if (-not $isModuleInstalled){
            Write-Host "You should install powershell module '$module'" -ForegroundColor DarkYellow
            $allModulesInstalled = $false
        }

    }
    
    # Return values
    return $allModulesInstalled
   
}

# Title
# isPowershellIse
#
# Params
# silent. boolean. indicates if function should print (false) or not (true).
#
# Description
# Checks if script is being run on powershell ise and suggests fixes if it is not.
#
# Return
# Boolean. True if it is run on ISE. Else, false.
# 
function isPowershellIse {
    param(
        [Parameter(Mandatory=$true)]
        [bool]$silent
    )

    if (-not ($silent)){
        Write-Host "Checking if tool is being run on powershell ise..." -ForegroundColor DarkCyan
    }

    if ($PSISE){
        Write-Host "You should not run tool from powershell ise" -ForegroundColor DarkYellow
        return $true
    }

    return $false
}

# Title
# SetPowershellExecutionPolicy
#
# Params
# None
#
# Description
# Sets powershell execution policy to unrestricted
#
# Return
# Boolean. True if powershell execution policy was successfully set to unrestricted. Else, false.
# 
function SetPowershellExecutionPolicy {

    # Attempts to set powershell execution policy
    try{
        Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted -Force
    }
    catch{
        Write-Host "Set-executionPolicy failed with error: $a_.Exception.Message" -ForegroundColor DarkRed
        return $false
    }

    return $true
}

# Title
# InstallRequiredPowershellModules
#
# Params
# None
#
# Description
# If user agrees, function installs powershell modules required by MD365Collector
#
# Return
# boolean. True if success, else false.
#
function InstallRequiredPowershellModules {

    # Loops over required powershell modules
    foreach ($module in $powershellModulesRequiredByTool) {
        
        #check if module is installed
        $isModuleInstalled = Get-Module -ListAvailable -Name $module

        #if it is not installed, attempts installation
        if (-not $isModuleInstalled){

            try {
                Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
            }
            catch {
                Write-Host "Module <$module> installation failed with error: $a_.Exception.Message" -ForegroundColor DarkRed
                return $false
            }

        }

    }

    return $true
}


# Title
# ImportRequiredModules
#
# Params
# None
#
# Description
# Loads powershell modules
#
# Return
# Boolean. True if import action went success. Else, false.
# 
function ImportRequiredModules {

    # Loops over all required powershell modules
    foreach ($module in $powershellModulesRequiredByTool){

        # attempt to load them
        try{
            Import-Module $module   
        }
        catch{
            Write-Host "Import-Module '$module' failed." -ForegroundColor DarkRed
            return $false
        }
    }

    return $true
}

# Title
# AuthenticateAsUserInExchangeOnline
#
# Params
# None
#
# Description
# Authenticates against Exchange Online PowerShell
#
# Return
# Boolean. True if authentication is success. Else, false.
# 
function AuthenticateAsUserInExchangeOnline {

    try{
        Connect-ExchangeOnline
    }
    catch{
        Write-Host "Authentication failed." -ForegroundColor DarkRed
        return $false
    }

    return $true
}

