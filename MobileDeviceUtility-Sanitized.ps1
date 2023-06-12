
param(
    [Parameter(Position = 0, Mandatory = $true, HelpMessage = "Enter filename with only a list of email addresses")]
    [string]$EmailListFile, 
    [Parameter(Mandatory = $false)]
    [bool]$AddUserToJamf = $false,
    [Parameter(Mandatory = $false)]
    [bool]$CheckDevicesOnly = $true,
    [Parameter(Mandatory = $false)]
    [bool]$ClearStaleDevices = $false,
    [Parameter(Mandatory = $false)]
    [string]$LogFile = ".\MobileDeviceLog-" + (Get-Date -Format yyyyMMdd-HHmmss) + ".log"
    )

Function Remove-BlockedDevice {
    param(
     [Parameter(Mandatory)]
     [string]$Identity
    )
    $timeout = new-timespan -minutes 5
    $sw = [diagnostics.stopwatch]::StartNew()
    while ($sw.elapsed -lt $timeout){
       if (Get-EXOMobileDeviceStatistics -Identity $Identity | Where-Object {$_.Status -eq "AccountOnlyDeviceWipeSucceeded"}) {
          Remove-MobileDevice -Identity $Identity -Confirm:$false
          break
       }
       Start-Sleep -seconds 5
    }
 
    if (!(Get-EXOMobileDeviceStatistics -Identity $Identity -ErrorAction SilentlyContinue)) {
       return 0
    }
    else {
        return 1
    }
}
 
Function Add-UserToJamfGroup {
    param(
        [Parameter(Mandatory)]
        [string]$username
    )
 
    $base64creds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("basicauthcredentials"))
    $JamfBearerToken = (Invoke-RestMethod -Uri "https://companyname.jamfcloud.com/api/v1/auth/token" -Headers @{Authorization = ("Basic $base64creds") } -Method Post).token
    
 
    $userBody = @(
        "<user_group>
            <user_additions>
                <user>
                    <username>$username</username>
                </user>
            </user_additions>
        </user_group>"
    )
 
    Invoke-WebRequest -Uri 'https://companyname.jamfcloud.com/JSSResource/usergroups/id/6' -Method PUT -body $userBody -Headers @{Authorization = "Bearer $JamfBearerToken" }
 
}

$AllEASDevices = @()
$AllStaleDevices = @()
$AllEASFile = ".\EASDevices-" + (Get-Date -Format yyyyMMdd-HHmm) + ".csv"
$AllStaleFile = ".\StaleDevices-" + (Get-Date -Format yyyyMMdd-HHmm) + ".csv"

Start-Transcript -Path $LogFile
Clear-Host
$Connections = Get-ConnectionInformation
If (!($Connections.ConnectionUri -match 'outlook\.office365\.com') ){
    Connect-ExchangeOnline -ShowBanner:$False
}

if (test-path $EmailListFile){
    $emails = Get-Content $EmailListFile
}
else {
    Write-Host "File not found. Stopping process." -ForegroundColor Red
    Exit
}

Write-Host `r`n
foreach ($email in $Emails){
    Write-Host "Processing $email..." -ForegroundColor Green    
    $username = (get-exomailbox $email).alias
    $Displayname = (get-exomailbox $email).DisplayName
    $devices = Get-EXOMobileDeviceStatistics -Mailbox $email -ErrorAction SilentlyContinue
    if ($devices) {
        $StaleDevices = $devices | Where-Object LastSuccessSync -lt (Get-Date).AddDays(-30)
        $EASDevices = $devices | Where-Object LastSuccessSync -gt (Get-Date).AddDays(-30) | Where-Object ClientType -eq "EAS"
        if ($StaleDevices){
            $i=0
            if ($ClearStaleDevices){
                Write-Host "`n`rThe following devices will be removed from $DisplayName's profile:"
            }
            else {
                Write-Host "`n`rThe following devices would be removed from $DisplayName's profile:"
            }
            $StaleDevices | ForEach-Object {
                $i++
                if ($Device.DeviceFriendlyName){
                    $FriendlyName
                }
                Write-Host -ForegroundColor Blue ("{0:D1}. {1} [Last Sync: {2}]" -f $i,$_.DeviceUserAgent,$_.LastSuccessSync)
            }
            $AllStaleDevices += $StaleDevices
        }
        if ($EASDevices) {
            $i=0
            if (!$CheckDevicesOnly){
                Write-Host "`n`rThe following devices will have ActiveSync corporate data wiped from $DisplayName's profile:"
            }
            else {
                Write-Host "`n`rThe following devices would have ActiveSync corporate data wiped from $DisplayName's profile:"
            }
            $EASDevices | ForEach-Object {
                $i++
                Write-Host -ForegroundColor Blue ("{0:D1}. {1} [Last Sync: {2}]" -f $i,$_.DeviceFriendlyName,$_.LastSuccessSync)
            }
            $AllEASDevices += $EASDevices
        }
        if (!$EASDevices -and !$StaleDevices) {
            write-host "No devices found that match criteria!" -ForegroundColor Yellow
        }
        foreach ($Device in $StaleDevices){
            if (($Device.LastSuccessSync -lt (Get-Date).AddDays(-30)) -and ($ClearStaleDevices)){
                try {
                    write-host `r`n"Removing device from user profile..." -ForegroundColor Yellow
                    write-host ($ident = $device.Identity) -ForegroundColor Blue
                    Remove-MobileDevice -Identity $ident -Confirm:$false
                }
                catch {
                    Write-Host ("Issue removing {0}" -f $_.Identity) -ForegroundColor Red
                }
            }
        }
        foreach ($Device in $EASDevices){
            if (($Device.LastSuccessSync -gt (Get-Date).AddDays(-30)) -and ($Device.ClientType -eq "EAS") -and (!$CheckDevicesOnly)){
                write-host `r`n"Wiping corporate profile from device..." -ForegroundColor Yellow
                write-host ($ident = $device.Identity) -ForegroundColor Blue
                Clear-MobileDevice -AccountOnly -Identity $ident
                $ExitCode = Remove-BlockedDevice -Identity $ident
                if ($ExitCode -eq 1){
                    Write-Host ("Issue removing {0}" -f $_.DeviceFriendlyName) -ForegroundColor Red
                }
            }

        }
        if ($AddUserToJamf){
            Add-UserToJamfGroup -username $username 
        }
        Write-Host `n`r
    } 
    else {
        write-host "No devices found!" -ForegroundColor Yellow
    }
}
$AllEASDevices | Export-CSV -Path $AllEASFile -NoTypeInformation
$AllStaleDevices | Export-CSV -Path $AllStaleFile -NoTypeInformation
Stop-Transcript | Out-Null