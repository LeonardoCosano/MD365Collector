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

# ToDo Remove force flag in production
Import-Module "$PSScriptRoot\utils.ps1" -Force

# Title
# GetAuditLogs
#
# Params
# User. String. Comma separated list of user principal names to be investigated.
# StartDate. String. Optional. Starting date from the log investigation timeframe
# EndDate. String. Optional. Ending date from the log investigation timeframe
#
# Description
# It collects audit log events from microsoft office 365. Right now you can only use following filters:
# username, timeframe.
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
        if($results.Count -cge 4999){
            Write-Host "Watch out! Only 5k results were returned, if you need an exhaustive list please reduce the searching timeframe" -ForegroundColor DarkYellow
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
    $results | Select-Object * | Export-Csv "$outputsFolderPath\O365AuditLogs.csv" -NoTypeInformation
    Write-Host "AuditLog search ended successfully. Saving at $outputsFolderPath\O365AuditLogs.csv" -ForegroundColor DarkGreen

    return $true   
}


# Title
# GetExpandedAuditLogsFile
#
# Params
# fileName. String. local file containing microsoft office 365's auditlogs.
#
# Description
# It will create a copy of the file, this new file will have more columns. New columns's names will be the keys from the json at original csv's colum "auditdata".
# This may be useful because for some auditlogs operations (mailitemsaccessed for example) the valuable part is stored in auditdata, inside a json, so it cant be easily parsed and filtered
#
# Return
# Boolean. True if audit logs were successfully retrieved. Else, false.
# File. CSV format file including auditlogs.
# 
function GetExpandedAuditLogsFile {
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

    ## 2.2. Save new headers, which are the key values of the json stored at auditdata colum.
    $newHeaders = @()
    foreach ($row in $csv){

        if ($row.AuditData) {
            $jsonAuditData = $row.AuditData | ConvertFrom-Json
            $newHeaders += $jsonAuditData.PSObject.Properties.Name
            #ToDo. Handle json properly. the auditlog includes nested jsons which are not properly dropped into the new file.
         }

     }

    $newHeaders = $newHeaders | Sort-Object -Unique

    # 3. Create a list of desired headers (columns from the new csv)
    $desiredHeaders = $originalHeaders
    foreach ($header in $newHeaders){
       $desiredHeaders += "${header}Json"
    }

    # 4. Create a csv file containing only the first header row with the desired headers
    $outputPath = [System.IO.Path]::ChangeExtension($InputCsvPath, "expanded.csv")
    $headerRow = ($desiredHeaders -join ",")
    $headerRow | Out-File -FilePath $outputPath -Encoding UTF8

    $counter=1
    $total=$csv.length

    Write-Host 'Creating a new file with expanded properties. Hang a moment, this may take a few minutes' -ForegroundColor DarkCyan
    # 5. Append rows to new csvFile
    foreach ($row in $csv){
        $counter = $counter + 1

        if ($counter -eq [int]($total/4)){ 
            write-host "25% rows expanded" -ForegroundColor DarkYellow
        }
        if ($counter -eq [int]($total/2)){
            write-host "50% rows expanded" -ForegroundColor DarkYellow  
        }
        if ($counter -eq [int]($total/4)*3){
            write-host "75% rows expanded" -ForegroundColor DarkYellow  
        }
        
        ## 5.1. Defines the new set of data to be pushed to csv
        $newRowData = @{}

        ## 5.2. Adds original data
        foreach ($header in $originalHeaders){
            $newRowData["$header"] = $row.$header
        }

        ## 5.3. Add new headers data
        $rowJson = $row.AuditData | ConvertFrom-Json
        foreach ($header in $newHeaders){
            $headerName = $header + "Json"

            if ($rowJson.$header -is [System.Object[]]){

               $newRowData["$headerName"] = $($rowJson.$header | ConvertTo-Json)#Out-String)                 

            } else {

                $newRowData["$headerName"] = $rowJson.$header
            
            }

        }
        
        ##5.4. Push the data to row
        $newRowDataObject = [PSCustomObject]$newRowData
        $newRowDataObject | Export-Csv $outputPath -NoTypeInformation -Append -Force
            
    }
    
    Write-Host 'Extraction is ready. You can now inspect results on $($outputPath)' -ForegroundColor DarkGreen   
}


# Title
# GetReadMailsFromAuditLogsFile
#
# Params
# fileName. String. csv file containing microsoft defender for office 365's auditlogs.
# showErrors. boolean. indicates if function should print (true) or not (false) details about mails not found.
# daysBack. int. helps setting the time window for searching email details. I recommend starting by default and continue adding 10, 20, 30...10*x if some details is missing and you really need it.
#
# Description
# It will read the original CSV file containing auditlogs from microsoft defender and create a new csv which will contain data about the emails.
#
# Return
# Boolean. True if mail data was successfully retrieved. Else, false.
# File. CSV format file including auditlogs.
# 
function GetReadMailsFromAuditLogsFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$fileName,
        [Parameter(Mandatory=$false)]
        [switch]$showErrors,
        [Parameter(Mandatory=$false)]
        [int]$daysBack=0
    )

    # 1. Check correct input parameter
    $InputCsvPath = Join-Path (Get-Location) $fileName
    if (-not (test-path $InputCsvPath)){
        write-host "Indicated file $filename was not found. We were looking on folder $(Get-location)" -ForegroundColor DarkRed
        return $false
    }

    #2. Read every event from auditlogs to save data already stored in auditlogs ("AADSessionId", "ClientIp", "EmailReadId", "EmailReadInmutableId", "EmailReadInternetMessageId")
    $AADSessionIdentifiers = @()
    $ClientIps =@()
    $emailIdentifier = @()
    $emailIdentifierInmutable = @()
    $emailIdentifierInternet = @()

    $csv = Import-Csv $InputCsvPath
    foreach ($event in $csv){
        
        if ($event.Operations -ne "MailItemsAccessed"){
            continue #skip this event 
        }

        # Each mailItemAccess event contains a json on AuditData field
        $mailDetailsJson = $event.AuditData | ConvertFrom-Json
        # Each read mails is inside a folder
        foreach ($readFolder in $mailDetailsJson.Folders){
            foreach($readMail in $readFolder.FolderItems){
                #Skip repeated values of the internet message id
                if ($emailIdentifierInternet -notcontains $readMail.InternetMessageId){
                    $AADSessionIdentifiers += $mailDetailsJson.AppAccessContext.AADSessionId
                    $ClientIps += $mailDetailsJson.ClientIPAddress
                    $emailIdentifier += $readMail.Id
                    $emailIdentifierInmutable += $readMail.ImmutableId
                    $emailIdentifierInternet += $readMail.InternetMessageId   
                }                            
            }
        }
    }   

    #3. Get read emails data which is not stored on the auditlogs ("EmailReadInternetMessageId", "EmailSubject", "EmailFrom"), based on internetmessageid got in prev step.
    $emailSubject = @()
    $emailSender = @()
    $emailIdentifierInternetParsed = $($emailIdentifierInternet -join ",")
    $filterEndDate = (Get-Date).AddDays(-$($daysBack)).ToString("dd/MM/yyyy")
    $filterStartDate = ((Get-Date).AddDays(-10-$($daysback))).ToString("dd/MM/yyyy")

    Write-Host "Searching email details on exchange from $($filterStartDate) to $($filterEndDate)" -ForegroundColor DarkCyan
    $AllEmailDetails = Get-MessageTraceV2 -MessageId "$($emailIdentifierInternetParsed)" -EndDate $($filterEndDate) -StartDate $($filterStartDate) | select-object MessageId, SenderAddress, Subject


    #4. Store data from both 2 and 3 step into csv
    $outputCsv = [System.IO.Path]::ChangeExtension($InputCsvPath, "AccessedEmails.csv")
    $headers = @("AADSessionId", "ClientIp", "EmailReadId", "EmailReadInmutableId", "EmailReadInternetMessageId", "EmailSubject", "EmailFrom")
    if (-not (Test-Path $outputCsv)) {
        $($headers -join ",") | Out-File -FilePath $outputCsv -Encoding UTF8
    }


    $counter = 1
    $mailsNotFound = 0
    $total = $emailIdentifierInternet.count

    foreach ($emailIndex in 1..$($emailIdentifierInternet.count)) {

        $counter = $counter + 1

        if ($counter -eq [int]($total/4)){ 
            write-host "25% mails searched" -ForegroundColor DarkYellow
        }
        if ($counter -eq [int]($total/2)){
            write-host "50% mails searched" -ForegroundColor DarkYellow  
        }
        if ($counter -eq [int]($total/4)*3){
            write-host "75% mails searched" -ForegroundColor DarkYellow  
        }

        $newEmailAADSession = $AADSessionIdentifiers[$($emailindex)]
        $newEmailClientIp = $ClientIps[$($emailindex)]
        $newEmailIdentifier = $emailIdentifier[$($emailindex)]
        $newEmailIdentifierInmutable = $emailIdentifierInmutable[$($emailindex)]
        $newEmailIdentifierInternet = $emailIdentifierInternet[$($emailindex)]
        $newEmailDetails = $AllEmailDetails | where-object { $_.MessageId -match $newEmailIdentifierInternet} | Select-Object -First 1
        $newEmailSubject = $newEmailDetails.Subject
        $newEmailSender = $newEmailDetails.SenderAddress

 
        if ($newEmailSubject -eq $null){
            if ($showErrors -eq $true){
                write-host "No data subject neither sender found for $newEmailIdentifierInternet" -ForegroundColor DarkYellow
            }
            $mailsNotFound = $mailsNotFound + 1
            $newEmailSubject = "Details not found."
            $newEmailSender = "Details not found."
        }
        
        $dataToCsv = [PSCustomObject]@{
            AADSessionId = $newEmailAADSession
            ClientIp = $newEmailClientIp
            EmailReadId = $newEmailIdentifier
            EmailReadInmutableId = $newEmailIdentifierInmutable
            EmailReadInternetMessageId = $newEmailIdentifierInternet
            EmailSubject = $newEmailSubject
            EmailFrom = $newEmailSender
        }

        $dataToCsv | Export-Csv $outputCsv -NoTypeInformation -Append  
        

    }

    #5. Save results
    Write-Host "Read mails extraction is ready. You can now inspect results on $($outputCsv)" -ForegroundColor DarkGreen
    Write-host "$($mailsNotFound) of $($total) mails were not found. You may want to use -daysBack" -ForegroundColor DarkYellow

}


# Title
# GetAuditLogActivities
#
# Params
# fileName. String. csv file containing microsoft defender for office 365's auditlogs.
#
# Description
# It will read the original CSV file containing auditlogs from microsoft defender and print the activities detected and the times it were logged.
#
# Return
# Boolean. True if stats were successfully calculated. Else, false.
# 
function GetAuditLogActivities {
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

    # 2. Load CSV
    $auditLogsCsv = Import-Csv -Path $InputCsvPath

    # 3. Parse data to get operations
    $operations = $auditLogsCsv | group-object -property Operations | Sort-Object -property Count -Descending

    # 4. Save operations into csv
    $stats = $operations | ForEach-Object {
        [PSCustomObject]@{
            OperationName = $_.Name
            Count     = $_.Count
        }
    }
    $outputCsvPath = [System.IO.Path]::ChangeExtension($InputCsvPath, "ActivitiesCount.csv")
    $stats | Export-Csv -Path $outputCsvPath -NoTypeInformation -Force

    # 5. Print stats
    Write-Host "Activity count was completed successfuly. You can now inspect results on $($outputCsvPath)" -ForegroundColor DarkGreen
    foreach ($activity in $stats){
        write-host ("{0,-45}" -f "$($activity.OperationName)") -NoNewline -ForegroundColor DarkYellow
        write-host " was detected " -NoNewline
        write-host ("{0,-5}" -f "`t$($activity.Count) ") -NoNewline -ForegroundColor DarkYellow
        write-host "`ttimes"
    }
  
}
