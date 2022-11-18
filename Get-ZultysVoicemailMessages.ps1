Write-Host "Welcome to the Zultys Voicemail mass downloader. `n This will loop through all messages in your mailbox and save them to a folder of your choice." -ForegroundColor Cyan
$SaveFolderPath = Read-Host "Please enter the path (no trailing \) where to save your voicemails"
Write-Host "Please Note: You will need to use the browser developer tools to retrieve your session GUID. `n Please avoid opening your voicemail area once you've signed into WebZAC" -ForegroundColor Yellow
Write-Host "Open Dev Tools, login to webzac, and pull the session GUID that is inside the get request"
$SessionGUID = Read-Host "Please enter the session guid you've copied here. Do not include anything but the guid"
$zultysUrl = "https://$(Read-Host 'Enter the hostname of your phone system.')"
$ZacAPIUrl = "$zultysUrl/newapi/?session=$($SessionGUID)&command="
$GetVoicemailsCommand = "cli_get_mailbox_media&page=1&limit=100&mediaType=VoiceMail&groupId=undefined"


$VoicemailBox = Invoke-RestMethod -Method Get -Uri "$($ZacAPiUrl)$($GetVoicemailsCommand)"

if ($VoicemailBox) {
    Write-Host "Found $($VoicemailBox.mailboxMedia.count) messages. `n Would you like to download them all to $SaveFolderPath" -ForegroundColor Cyan
    $ConfirmDownload = Read-Host "Y/n"

    if ($ConfirmDownload -in 'y','Y') {

        foreach ($Voicemail in $VoicemailBox.mailboxMedia) {
            $VoicemailIDtoDownload = $Voicemail.mediaId
            $DownloadMessagesCommand = "cli_get_media_file&mediaId=$VoicemailIDtoDownload&mediaType=VoiceMail"
            $VoicemailDate = (Get-Date ([datetime]'1970-01-01 00:00:00').addseconds($Voicemail.timestamp) -Format yyyyMMddHHmm)
            $OutFileName = "VM-$($Voicemail.callerName)-$VoicemailDate.wav"
            $OutFileMetadata = "VM-$($Voicemail.callerName)-$VoicemailDate.txt"

            $Voicemail | Out-File "$SaveFolderPath\$OutFileMetadata"
            Invoke-WebRequest -Method GET -Uri "$($ZacAPIUrl)$($DownloadMessagesCommand)" -OutFile "$SaveFolderPath\$OutFileName"

            $DownloadedFiles = Get-ChildItem -Path $SaveFolderPath -Filter *.wav
            if ($DownloadedFiles.count -eq $VoicemailBox.mailboxMedia.count) {
                Write-Host "Downloaded all files successfully" -ForegroundColor Blue
            }
            else {
                Write-Host "Failed to download all messages. Only $($DownloadedFiles.count) were successful" -ForegroundColor Red
            }
            

        }

    }
}
