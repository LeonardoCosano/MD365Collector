# Title 
# Global variable powershellModulesRequiredByTool
# 
# Description
# Variable containing powershell module names which are required by MD365Collector to work.
# 
$script:powershellModulesRequiredByTool = @(
    'ExchangeOnlineManagement'
    #'Az'
)


# Title
# CheckPowershellModulesAvailability
#
# Params
# None
#
# Description
# Checks if the machine executing MD365Collector has required powershell modules installed
#
# Return
# Boolean. True if machine has all required powershell modules installed, else false.
# 
function CheckPowershellModulesAvailability {

    # Iterate over list of required modules
    $allModulesInstalled = $true
    foreach ($module in $powershellModulesRequiredByTool) {

        #check if module is installed
        $isModuleInstalled = Get-Module -ListAvailable -Name $module

        # If just one is not installed, then we cant stop
        if (-not $isModuleInstalled){
            $allModulesInstalled = $false
            break
        }
    }
    
    # Return values
    if (-not $allModulesInstalled)  {
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

    # Requests user agreement
    $UserAgreement = Read-Host "Do you agree to install powershell modules which are required by the tool? (Y/n)"
    if ($UserAgreement -notin @('Y','y')) {
        Write-Host "Aborting MD365Collector. Install following modules to successfuly execute:" -ForegroundColor 
    }

    # Loops over required powershell modules
    foreach ($module in $powershellModulesRequiredByTool) {
        
        #check if module is installed
        $isModuleInstalled = Get-Module -ListAvailable -Name $module

        #if it is not installed, tries installation
        if (-not $isModuleInstalled){
            Write-Host "Module <$module> is missing." -ForegroundColor Yellow
            $installationStatus = InstallModuleIfUserAgrees -Modulename $module -Agreement $UserAgreement

            #if installation fails, tool exits
            if ($installationStatus == $false) {
                Write-Host "Module <$module> was not installed but it is required." -ForegroundColor Red
                Write-Host "Install module <$module> to continue using MD365Collector" -ForegroundColor Red
                return $false
            }         
        }
    }

    return $true

}

# Title
# InstallModuleIfUserAgrees
#
# Params
# ModuleName. String. The name of the module to install.
# Agreement. String (Y/y/N/n). The value of the user agreement
#
# Description
# If user agrees, function installs powershell modules required by MD365Collector. Else, does nothing.
#
# Return
# boolean. True if installation success, else false.
#
function InstallModuleIfUserAgrees {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ModuleName,

        [Parameter(Mandatory=$true)]
        [string]$Agreement
    )

    # If user did not grant permissions, then installation is cancelled
    if ($Agreement -notin @('Y','y')){
        return $false
    }

    # Attempts installation
    try {
        Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
    }
    catch {
        return $false
    }

    return $true

}

# Title
# CheckPowershellExecutionPolicy
#
# Params
# None
#
# Description
# Checks if the powershell execution policy is unrestricted or bypass
#
# Return
# Boolean. True if powershell execution policy is unrestricted or bypass. Else, false.
# 
function CheckPowershellExecutionPolicy {
    $policy = Get-ExecutionPolicy -Scope CurrentUser
    if ($policy -notin @('Unrestricted ','Bypass')){
        return $false
    }
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

    # Requests user agreement
    $UserAgreement = Read-Host "Do you agree to set powershell execution policy to unrestricted? (Y/n)"

    #Exits if it is not granted
    if ($UserAgreement -notin @('Y','y')) {
        Write-Host "Aborting MD365Collector. Unrestricted execution policy is required for the successful execution of the tool." -ForegroundColor Yellow
        return $false
    }

    # Attempts to set powershell execution policy
    try{
        Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted -Force
    }
    catch{
        Write-Host "Set-executionPolicy failed." -ForegroundColor Red
        return $false
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
            Write-Host "Import-Module '$module' failed." -ForegroundColor Red
            return $false
        }
    }

    return $true
}

# Title
# AuthenticateInExchangeOnline
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
function AuthenticateInExchangeOnline {

    $UPN = Read-Host "Introduce UserPrincipalName of the account that will be used to authenticate"

    try{
        Connect-ExchangeOnline -UserPrincipalName $UPN
    }
    catch{
        Write-Host "Authentication failed." -ForegroundColor Red
        return $false
    }

    return $true
}

# Title
# CheckPowershellIse
#
# Params
# None
#
# Description
# Checks if script is being run on powershell ise
#
# Return
# Boolean. True if it is run on ISE. Else, false.
# 
function CheckPowershellIse {

    if ($PSISE){
         Write-Host "Do not" -ForegroundColor Yellow
        return $true
    }

    return $false
}