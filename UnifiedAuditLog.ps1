# Title
# checkUTCDate
#
# Params
# Date. String. Date value in unkown format
#
# Description
# It confirms wether variable is datetime in format UTC+0.
#
# Return
# Boolean. True if date is UTC+0. Else, false.
# 
function checkUTCDate {
    param(
        [Parameter(Mandatory=$true)]
        [DateTime]$Date
    )

    # Checks wether parameter is in a valid date variable
    if (-not ($date -is [DateTime])) {
        Write-Host "Date parameter is not a valid datetime variable" -ForegroundColor Red
        return $false 
    } 

    # Catches exceptions due to access Kind property of a date variable.
    try{
        $date.Kind
    }
    catch{
        Write-Host "Error while inspecting date parameter date format kind property $a_.Exception.Message" -ForegroundColor Red  
        return $false
    }

    # Checks wether date variable is in UTC
    if ($date.Kind -eq [System.DateTimeKind]::Utc) {
        return $true
    }

    # If date variable is not UTC+0, return false. 
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
# It searches audit log events from the username indicated.
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

    # If custom values for date parameters are not provided, set a default value
    if (-not ($PSBoundParameters.ContainsKey('end')) -or -not ($end)) {
        $end = (Get-Date).ToUniversalTime()
    }        
    if (-not ($PSBoundParameters.ContainsKey('start')) -or -not ($start)) {
        $start = $end.AddDays(-1)
    }

    # If date parameteres are provided (left statement of the and condition) and its format is wrong (right statement of and condition) set correct value
    if ($PSBoundParameters.ContainsKey('end') -and -not (checkUTCDate -date $end)){
        Write-Host "Your end date parameter is being treated as UTC your Localtime" -ForegroundColor Yellow
        $end = $end.ToUniversalTime()
    }
    if ($PSBoundParameters.ContainsKey('start') -and -not (checkUTCDate -date $start)){
        Write-Host "Your start date parameter is being treated as UTC your Localtime" -ForegroundColor Yellow
        $start = $start.ToUniversalTime()      
    }


    #Ãttempt to search AuditLogs
    try{
        Search-UnifiedAuditLog -EndDate $end -StartDate $start -UserIds $user    
    }
    catch{
        Write-Host "Error while Search-UnifiedAuditLog execution: $_.Exception.Message" -ForegroundColor Red
        return $false          
    }

    return $true   
}