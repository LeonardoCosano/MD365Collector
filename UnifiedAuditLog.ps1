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
  Purpose: Features & functions to interact with AuditLogs feature from Microsoft Defender for office 365.
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
# GetAuditLogs
#
# Params
# UserName. String. Comma separated list of user principal names to be investigated.
# StartDate. String. Optional. Starting date from the log investigation timeframe
# EndDate. String. Optional. Ending date from the log investigation timeframe
#
# Description
# It searches audit log events from the username indicated during the timeframe established. It also saves output into a output folder
#
# Return
# Boolean. True if audit logs were successfully retrieved. Else, false.
# File. CSV format file including auditlogs.
# 
function GetAuditLogs {
    param(
        [Parameter(Mandatory=$true)]
        [string]$user,

        [Parameter(Mandatory=$false)]
        [DateTime]$start,

        [Parameter(Mandatory=$false)]
        [DateTime]$end
    )

    # 1. Handle parameters
    ## 1.1. If custom values for date parameters are not provided, set a default value
    if (-not ($PSBoundParameters.ContainsKey('end')) -or -not ($end)) {
        $end = (Get-Date).ToUniversalTime()
    }        
    if (-not ($PSBoundParameters.ContainsKey('start')) -or -not ($start)) {
        $start = $end.AddDays(-1)
    }

    ## 1.2.  If date parameteres are provided (left statement of the and condition) and its format is wrong (right statement of and condition) set correct value
    if ($PSBoundParameters.ContainsKey('end') -and -not (isDateUTC -Date $end)){
        $end = $end.ToUniversalTime()
        Write-Host "Your end date parameter value has been adapted to UTC+0: $end" -ForegroundColor DarkYellow

    }
    if ($PSBoundParameters.ContainsKey('start') -and -not (isDateUTC -Date $start)){
        $start = $start.ToUniversalTime()    
        Write-Host "Your start date parameter value has been adapted to UTC+0: $start" -ForegroundColor DarkYellow
    }

    # 2. Start the search
    Write-Host "Searching audit logs for $user since $start to $end" -ForegroundColor DarkCyan

    try{
        # Create a uniqueIdentifier, the epoch time
        $sessionID = [int][double]((New-TimeSpan -Start (Get-Date "1970-01-01T00:00:00Z") -End (Get-Date).ToUniversalTime()).TotalSeconds)
        $results = Search-UnifiedAuditLog -EndDate $end -StartDate $start -UserIds $user -HighCompleteness -ResultSize 5000 -SessionCommand ReturnLargeSet -SessionId $sessionID -errorAction stop
    }
    catch{
        Write-Host "Error while Search-UnifiedAuditLog execution: $_.Exception.Message" -ForegroundColor DarkRed
        return $false          
    }

    # 3. Store results into outputs folder
    ## 3.1. Get outputs folder
    $outputsFolderPath = "outputs/" + (Get-Date).ToString("yyyy.MM.dd")+"-"+$user
    if (-not (Test-Path $outputsFolderPath)) {
        New-Item -Path $outputsFolderPath -ItemType Directory | Out-Null
    }
    ## 3.2. Save results
    $results | Select-Object * | Export-Csv "$outputsFolderPath\O365AuditLogs.csv" -NoTypeInformation
    Write-Host "AuditLog search ended successfully. Saving at $outputsFolderPath\O365AuditLogs.csv" -ForegroundColor DarkGreen

    return $true   
}