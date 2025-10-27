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
# UserName. String. Comma separated list of user principal names to be investigated.
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
# ExtractAuditData
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
    $outputCsv = [System.IO.Path]::ChangeExtension($InputCsvPath, "readMails.csv")
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

        # Each mailItemAccess contains at least 1 mail read, for each one, we get properties
        $mailDetailsJson = $event.AuditData | ConvertFrom-Json
        foreach ($folder in $mailDetailsJson.Folders){
            foreach($item in $folder.FolderItems){
                
                # this call here gets properties
                $readEmailDetails = Get-MessageTraceV2 -MessageId "$($item.InternetMessageId)"

                # If no properties are found related to the internetmessageid, this data columns are not filled
                if (-not ($readEmailDetails)){
                    write-host "No mail has been found with InternetMessageId $($item.InternetMessageId)" -ForegroundColor DarkYellow
                    $emailSubject = ""
                    $emailTo = ""
                    $emailFrom = ""
                } else {

                    $emailSubject = $readEmailDetails[0].subject
                    $emailTo = $readEmailDetails[0].RecipientAddress
                    $emailFrom = $readEmailDetails[0].SenderAddress
                
                }

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
                    EmailSubject = $emailSubject
                    EmailTo =$emailTo
                    EmailFrom = $emailFrom
                }

                $dataToCsv | Export-Csv $outputCsv -NoTypeInformation -Append        
                       
            }
        }                
    }

    #4. Save results
    Write-Host 'Read mails extraction is ready. You can now inspect results on $outputPath' -ForegroundColor DarkGreen


}