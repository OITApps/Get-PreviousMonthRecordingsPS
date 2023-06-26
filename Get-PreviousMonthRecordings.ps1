# Instructions
# Create a folder for call recordings, and save this file there
# Create a task schedule event to run this powershell once per month. 
# It will create a folder for the previous calendar month, download all recordings, and create a manifest CSV file

# Credentials for connecting to switch
# Must be a user with Office Manager permissions or abvove
$fqdn = ""
$domain = ""
$clientID = ""
$clientSecret = ""
$userName = ""
$password = ""

$debug = 0 #set to 1 to output to ps shell, set to 0 to disable shell output

#region Helper Functions

# logic for enabling/disabling debug in the script. 
if ($debug -eq 1) {
    $output = "Write-Output"
} else {
    $output = { param($args) }
}

## Trap any errors
# trap [Net.WebException] { continue; }
#Add Web Assembly for URL encoding
Add-Type -AssemblyName System.Web
# Authenticate against switch
Function Get-Token {
    ## Helper function to get an access token. Required to perform calls against the API
    ## Scopes: Any
    $tokenURL = "https://" + $fqdn + "/ns-api/oauth2/token/?grant_type=password&client_id=" + $clientID + "&client_secret=" + $clientSecret + "&username=" + $userName + "&password=" + $password

    $response = Invoke-RestMethod $tokenURL 
    $currentDate = Get-Date

    $Global:apiToken = New-Object -TypeName psobject
    $apiToken | Add-Member NoteProperty -Name accesstoken -Value $response.access_token
    $apiToken | Add-Member NoteProperty -Name expiration -Value $currentDate.AddSeconds(3600)
}

Function Invoke-NSRequest {
    ## Helper function to place API calls
    ## Scopes: Any
    param (
        [Parameter(Mandatory = $true)][Hashtable]$load,
        [Parameter(Mandatory = $false)][String]$type
    )
    # Check if payload submitted
    if (!$load) {
        Write-Host -ForegroundColor Red "Invalid or missing payload. Killing application"
        exit
    }
    # NS token expires in 1 hour. Check if token is still valid. If not, request a new one
    if ((!$apitoken) -or ((Get-Date) -lt $apitoken.expiration)) {
        Get-Token
    }

    # Check if request is POST or GET. Set GET by default
    if (!$type) { $type = "GET" }

    # Add format descriptor in case it's missing
    if (!$load.format) { $load.add('format', 'json') }

    # Set headers
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", 'Bearer ' + $apitoken.accesstoken)

    # Set request URL
    $requrl = "https://" + $fqdn + "/ns-api/"

    $response = Invoke-RestMethod $requrl -Headers $headers -Method $type -Body $load
    return $response
}
Function Convert-EpochDate($epoch){
    [datetime]$origin = '1970-01-01 00:00:00'
    $res = get-date $origin.AddSeconds($epoch) -Format d
    return $res
}
Function Convert-EpochTime($epoch){
    [datetime]$origin = '1970-01-01 00:00:00'
    $res = get-date $origin.AddSeconds($epoch) -Format T
    return $res
}

#endregion

function Create-FileObject {
    param(
        [string]$Date,
        [string]$Time,
        [string]$From,
        [string]$To,
        [string]$Filename,
        [string]$Duration
    )

    $file = [PSCustomObject]@{
        Date = $Date
        Time = $Time
        From = $From
        To = $To
        Duration = $Duration
        Filename = $Filename
    }
    $filelist += $file  # Add the created file object to the $filelist array
    return $file

}

#Set start and end date to last calendar month
#Format dates to 2019-01-01 00:00:00
function Format-NSTime($data){
    $data = Get-Date -UFormat "%Y-%m-%d %H:%M:%S"
    return $data
}
$startMonth = (Get-Date -Day 1).Date.AddMonths(-1) | Get-Date -UFormat "%Y-%m"
$startDate = (Get-Date -Day 1).Date.AddMonths(-1)
$startDate = Get-Date $startDate -UFormat "%Y-%m-%d %H:%M:%S"
$endDate = (Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0).AddSeconds(-1)
$endDate = Get-Date $endDate -UFormat "%Y-%m-%d %H:%M:%S"

$payload = @{
    object     = 'cdr2'
    action     = 'read'
    start_date = $startDate
    end_date   = $endDate
    domain     = $domain
    raw        = 'no'
    limit      = '9999999'
    format     = 'json'
}  
Try {
    $cdrs = Invoke-NSRequest $payload 
    $cdrs = $cdrs.CdrR | Select-Object -Property orig_callid, orig_to_user, term_callid, orig_from_user, time_start, duration
    # Invoke-Output $cdrs
    & $output $cdrs

    # Create folder for call recordings
    $folderName = $psscriptroot + "\" + $startMonth
    if (!(Test-Path $folderName)) {
        New-Item -Path $folderName -ItemType Directory
    }
    $filelist = @()

    foreach ($cdr in $cdrs) {
        $payload = @{
            object      = 'recording'
            action      = 'read'
            orig_callid = $cdr.orig_callid
            term_callid = $cdr.term_callid
            domain      = $domain
            limit       = '999999'
        }
        $retryCount = 1
        $retryLimit = 3
        $fileDownloaded = $false
        while (!$fileDownloaded -and $retryCount -lt $retryLimit) {
        try {
        $call = Invoke-NSRequest $payload 

        $fileName =  $cdr.term_callid + ".wav"
        $newPath = $folderName + "\" + $fileName

        if ($call.url) { 
            # Download the file if a valid URL exists
            $call.url | ForEach-Object { Invoke-WebRequest -Uri $_ -OutFile $newPath }

            $file = Create-FileObject -Date (Convert-EpochDate $cdr.time_start) -Time (Convert-EpochTime $cdr.time_start) `
                                     -From $cdr.orig_from_user -To $cdr.orig_to_user -Duration $cdr.duration -Filename $fileName
            & $output $file
            & $output "File downloaded: $newPath"
            $fileDownloaded = $true

            # Add the file object to the $filelist array
            $filelist += $file
        } else {
            $file = Create-FileObject -Date (Convert-EpochDate $cdr.time_start) -Time (Convert-EpochTime $cdr.time_start) `
                                     -From $cdr.orig_from_user -To $cdr.orig_to_user -Duration $cdr.duration -Filename $fileName

            & $output $file
            & $output "No Recording Found"
            break  # Exit the inner loop if "No Recording Found" or file downloaded
        }
    }
            catch {
                $errorMessage = $_.Exception.Message
                & $output "Error: $errorMessage"

                if ($retryCount -lt $retryLimit) {
                    $payload = @{
                        object      = 'recording'
                        action      = 'read'
                        orig_callid = $cdr.term_callid
                        term_callid = $cdr.orig_callid
                        domain      = $domain
                        limit       = '999999'
                    }
                    $fileName =  $cdr.orig_callid + ".wav"
                    $newPath = $folderName + "\" + $fileName
                    $file = Create-FileObject -Date (Convert-EpochDate $cdr.time_start) -Time (Convert-EpochTime $cdr.time_start) `
                                            -From $cdr.orig_from_user -To $cdr.orig_to_user -Duration $cdr.duration -Filename $fileName

                    & $output $file
                    & $output "Retrying..."
                } else {
                    $file = Create-FileObject -Date (Convert-EpochDate $cdr.time_start) -Time (Convert-EpochTime $cdr.time_start) `
                                            -From $cdr.orig_from_user -To $cdr.orig_to_user -Duration $cdr.duration -Filename $fileName

                    & $output $file
                    & $output "Retry limit reached. Skipping file."
                    $fileDownloaded = $true  # Set fileDownloaded to true to skip adding to $filelist
                    $file = $null  # Reset $file variable to null
                }
            }
        }
    }

    $manifestName = (Get-Date -Day 1).Date.AddMonths(-1) | Get-Date -UFormat "%Y%m"
    $filelist | ConvertTo-Csv -NoTypeInformation | Set-Content "$($folderName)\$($manifestName) Call History.csv"
}
Catch {
    $errorMessage = $_.Exception.Message
    & $output "Error: $errorMessage"
}
