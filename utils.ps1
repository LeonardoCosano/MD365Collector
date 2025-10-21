<#

 /$$      /$$ /$$$$$$$   /$$$$$$   /$$$$$$  /$$$$$$$   /$$$$$$            /$$ /$$                       /$$                        
| $$$    /$$$| $$__  $$ /$$__  $$ /$$__  $$| $$____/  /$$__  $$          | $$| $$                      | $$                        
| $$$$  /$$$$| $$  \ $$|__/  \ $$| $$  \__/| $$      | $$  \__/  /$$$$$$ | $$| $$  /$$$$$$   /$$$$$$$ /$$$$$$    /$$$$$$   /$$$$$$ 
| $$ $$/$$ $$| $$  | $$   /$$$$$/| $$$$$$$ | $$$$$$$ | $$       /$$__  $$| $$| $$ /$$__  $$ /$$_____/|_  $$_/   /$$__  $$ /$$__  $$
| $$  $$$| $$| $$  | $$  |___  $$| $$__  $$|_____  $$| $$      | $$  \ $$| $$| $$| $$$$$$$$| $$        | $$    | $$  \ $$| $$  \__/
| $$\  $ | $$| $$  | $$ /$$  \ $$| $$  \ $$ /$$  \ $$| $$    $$| $$  | $$| $$| $$| $$_____/| $$        | $$ /$$| $$  | $$| $$      
| $$ \/  | $$| $$$$$$$/|  $$$$$$/|  $$$$$$/|  $$$$$$/|  $$$$$$/|  $$$$$$/| $$| $$|  $$$$$$$|  $$$$$$$  |  $$$$/|  $$$$$$/| $$      
|__/     |__/|_______/  \______/  \______/  \______/  \______/  \______/ |__/|__/ \_______/ \_______/   \___/   \______/ |__/      
                                                                                                                                 

  Script: UnifiedAuditLog.ps1
  Author: Leonardo Cosano
  Purpose: Features & utli functions for MD365Collector such us outptusfolder creation, etc.
  Date: 2025.10.16

#>



# Title
# isDateUTC
#
# Params
# Date. String. Date value in unkown format
#
# Description
# It confirms wether variable is datetime in format UTC.
#
# Return
# Boolean. True if date is UTC. Else, false.
# 
function isDateUTC {
    param(
        [Parameter(Mandatory=$true)]
        [DateTime]$Date
    )

    # Checks whether parameter is in a valid date variable
    if (-not ($date -is [DateTime])) {
        Write-Host "Date parameter is not a valid datetime variable" -ForegroundColor DarkRed
        return $false 
    } 

    # Catches exceptions due to access Kind property of a date variable.
    try{
        # Checks wether date variable is in UTC
        $dateKind = $date.Kind
        if ($dateKind -eq [System.DateTimeKind]::Utc -or $dateKind -eq [System.DateTimeKind]::Local) {
            return $true
        }
    }
    catch{
        Write-Host "Error while inspecting date parameter date format kind property $a_.Exception.Message" -ForegroundColor DarkRed  
        return $false
    }

    # If date variable is not UTC, return false. 
    return $false

}

# Title
# isOutputFolderCreated
#
# Params
# path. String. the path of the folder to be checked.
#
# Description
# It will search if a folder exists.
#
# Return
# Boolean. True if folder already exists. Else, false.
# 
function isOutputFolderCreated {
    param(
        [Parameter(Mandatory=$true)]
        [string]$path
    )

    if (-not (Test-Path $path)) {
        return $false
    }

    return $true
}

# Title
# createOutputFolder
#
# Params
# path. String. the path of the folder to be created.
#
# Description
# It will create a folder.
#
# Return
# string. If folder is successfully created, true. Else, false.
# 
function createOutputFolder {
    param(
        [Parameter(Mandatory=$true)]
        [string]$path
    )

    try{
        New-Item -Path $path -ItemType Directory | Out-Null
    }
    catch{
        return $false
    }

    return $true
}


