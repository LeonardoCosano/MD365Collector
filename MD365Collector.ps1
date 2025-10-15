# ToDo Remove force flag in production
Import-Module "$PSScriptRoot\Prerequisites.ps1" -Force

# 1. Check execution prerequisites
## 1.1. Checks if powershell execution policy
if (-not (CheckPowershellExecutionPolicy)){
    $success = SetPowershellExecutionPolicy
}

if ($success = $false){
    exit 1
}

## 1.2. Checks if powershell modules required by the tool are installed on system
if (-not (CheckPowershellModulesAvailability)){
    $InstallationSuccess = InstallRequiredPowershellModules
}

if ($InstallationSuccess = $false){
    exit 1
}

## 1.3. Load required modules
if (-not (ImportRequiredModules)){
    exit 1
}

## 1.4 Check powershell ISE, which does not allow starting authentication process
if (CheckPowershellIse){
    exit 1
}

## 1.5. Authenticate
if (-not (AuthenticateInExchangeOnline)){
    exit 1
}

# ToDo
# 2...
# 3...
# 4...