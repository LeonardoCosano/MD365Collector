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

        #ToDo Handle edge cases where event logs count is above 5000.
        if($results.Count -cge 4900){
            Write-Host "Watch out! Only 5k results were returned, if you need an exhaustive list please access AuditLogs search tool" -ForegroundColor DarkYellow
        }

    }
    catch{
        Write-Host "Error while Search-UnifiedAuditLog execution: $_.Exception.Message" -ForegroundColor DarkRed
        return $false          
    }

    # 3. Store results into outputs folder
    ## 3.1. Get outputs folder
    $outputsFolderPath = "outputs/" + (Get-Date).ToString("yyyy.MM.dd") + "-" + $user
    if (-not (isOutputFolderCreated -path $outputsFolderPath)){
        createOutputFolder -path $outputsFolderPath
    }

    ## 3.2. Save results
    $results | Select-Object CreationDate, ResultIndex, UserIds, RecordType, Operations, AuditData  | Export-Csv "$outputsFolderPath\O365AuditLogs.csv" -NoTypeInformation
    Write-Host "AuditLog search ended successfully. Saving at $outputsFolderPath\O365AuditLogs.csv" -ForegroundColor DarkGreen

    return $true   
}

# Title
# ExtractAuditData
#
# Params
# fileName. String. file containing microsoft defender for office 365's auditlogs.
#
# Description
# It will create a copy of the file, this time with more columns. New columns will be the keys from the json at original csv's colum "auditdata".
#
# Return
# Boolean. True if audit logs were successfully retrieved. Else, false.
# File. CSV format file including auditlogs.
# 
function ExtractAuditData {
    param(
        [Parameter(Mandatory=$true)]
        [string]$fileName
    )

    # 1. Check correct input parameter
    $InputCsvPath = Join-Path (Get-Location) $fileName
    if (-not (test-path $InputCsvPath)){
        write-host "Indicated file $filename was not found. We were looking on folder $(Get-location)" -ForegroundColor DarkRed
        return $false
    }

    # 2. Obtain all csv headers

    $csv = Import-Csv $InputCsvPath 

    ## 2.1. Save original ones
    $originalHeaders = $csv[0].PSObject.Properties.Name

    ## 2.2. Save new headers, which are the key vaules of the json stored at auditdata colum.
    $newHeaders = @()
    foreach ($row in $csv){

        if ($row.AuditData) {
            $jsonAuditData = $row.AuditData | ConvertFrom-Json
            $newHeaders += $jsonAuditData.PSObject.Properties.Name
            #ToDo. Handle json properly. the auditlog includes nested jsons which are not properly dropped into the new file.
         }

     }


    $newHeaders = $newHeaders | Sort-Object -Unique

    # 3. Create a list of desired headers
    $desiredHeaders = $originalHeaders
    foreach ($header in $newHeaders){
       $desiredHeaders += "${header}Json"
    }

    # 4. Create a csv file containing only the first row, with the desired headers
    $outputPath = [System.IO.Path]::ChangeExtension($InputCsvPath, "_expanded.csv")
    $headerRow = ($desiredHeaders -join ",")
    $headerRow | Out-File -FilePath $outputPath -Encoding UTF8

    # 5. Append rows to new csvFile
    foreach ($row in $csv){
        
        ## 5.1. Defines the new set of data
        $newRowData = @{}

        ## 5.2. Adds original data
        foreach ($header in $originalHeaders){
            $newRowData["$header"] = $row.$header
        }

        ## 5.3. Add new headers data
        $rowJson = $row.AuditData | ConvertFrom-Json
        foreach ($header in $newHeaders){
            $headerName = $header + "Json"
            $newRowData["$headerName"] = $rowJson.$header
        }
        
        ##5.4. Push the data to row
        $newRowDataObject = [PSCustomObject]$newRowData
        $newRowDataObject | Export-Csv $outputPath -NoTypeInformation -Append -Force
            
    }
    
    Write-Host 'Extraction is ready. You can now inspect results on $outputPath' -ForegroundColor DarkGreen   
}

# Title
# ExtractReadMails
#
# Params
# fileName. String. file containing microsoft defender for office 365's auditlogs.
#
# Description
# It will read the original CSV auditlogs file from microsoft defender and output the number of read mails.
#
# Return
# Boolean. True if mail data was successfully retrieved. Else, false.
# File. CSV format file including auditlogs.
# 
function ExtractReadMails {
    param(
        [Parameter(Mandatory=$true)]
        [string]$fileName
    )

    # 1. Check correct input parameter
    $InputCsvPath = Join-Path (Get-Location) $fileName
    if (-not (test-path $InputCsvPath)){
        write-host "Indicated file $filename was not found. We were looking on folder $(Get-location)" -ForegroundColor DarkRed
        return $false
    }

    # 2. Create file for results
    $outputCsv = [System.IO.Path]::ChangeExtension($InputCsvPath, "_readMails.csv")
    if (-not (Test-Path $outputCsv)) {
        [PSCustomObject]@{
            TimeStamp = ""
            Operation = ""
            OperationProperty = ""
            AADSessionId = ""
            IssueTime =""
            ClientIp = ""
            ClientInfo = ""
            EmailReadId = ""
            EmailReadInmutableId = ""
            EmailReadInternetMessageId = ""
            EmailSubject = ""
            EmailTo = ""
            EmailFrom = ""
        } | Export-Csv $outputCsv -NoTypeInformation
    }

    #3. Read every event from auditlogs to save read emails
    $csv = Import-Csv $InputCsvPath 
    foreach ($event in $csv){
        
        # We are only looking for MailItemsAccessed events
        if ($event.Operations -ne "MailItemsAccessed"){
            continue
        }

        $mailDetailsJson = $event.AuditData | ConvertFrom-Json

        foreach ($folder in $mailDetailsJson.Folders){

            foreach($item in $folder.FolderItems){

                #emailData = Get-EXOMailboxMessage -Mailbox $mailDetailsJson.MailboxOwnerUPN -Filter "internetMessageId eq '$item.InternetMessageId'"

                $dataToCsv = [PSCustomObject]@{
                    TimeStamp = $mailDetailsJson.CreationTime
                    Operation = $mailDetailsJson.Operation
                    OperationProperty = $mailDetailsJson.OperationProperties.Name + $mailDetailsJson.OperationProperties.Value
                    AADSessionId = $mailDetailsJson.AppAccessContext.AADSessionId
                    IssueTime = $mailDetailsJson.AppAccessContext.IssuedAtTime
                    ClientIp = $mailDetailsJson.ClientIPAddress
                    ClientInfo = $mailDetailsJson.ClientInfoString
                    EmailReadId = $item.Id
                    EmailReadInmutableId = $item.ImmutableId
                    EmailReadInternetMessageId = $item.InternetMessageId
                    EmailSubject = "none"#emailData.subject
                    EmailTo ="none"#emailData.ToRecipients
                    EmailFrom = "none"#emailData.From
                }


                $dataToCsv | Export-Csv $outputCsv -NoTypeInformation -Append        
                       
            }

        }                

    }

    #4. Save results
    Write-Host 'Read mails extraction is ready. You can now inspect results on $outputPath' -ForegroundColor DarkGreen


}