<#
.SYNOPSIS
Send URLs and Files from an Outlook Email folder that contains unread messages to Cuckoo.

.DESCRIPTION
    PowerCuckoo
    Date: 8/14/2017
    Updated: 10/14/2019

    This script is currently gui/manually driven but can be automated by statically setting some variables.
    This release adds some automation to send artifacts to Cuckoo alongside the testing version to get an understanding of how it works. 
    The goal is to create a fully automated version that is liteweight and easy to use.

.EXAMPLE
./PowerCuckoo.ps1
Running the raw script will run the manual or automated version of the script depending on your answer.

./PowerCuckoo.ps1 -automated false
The will allow the script to run the manual version, prompting for your configuration settings and walk you through the process of submitting files.
This is good for testing!

./PowerCuckoo.ps1 -automated true
This will allow the script to run the automated version checking for the configuration in this script below.
If the settings are null or misconfigured it will alert you as such.

./PowerCuckoo.ps1 -automated true -duration 10
This will allow the script to run forever and will check for the specified amount of seconds before the scripts stops. 
Lower the seconds the more work the script has to do to continue to check for an unread item.
Note: This will check for unread messages every minute.

.NOTES
Currently works well for grabbing Unread messages from a folder of your choosing and sending them to your Cuckoo host. 
Try it out!

Lines 63 65 67 69 are used for statically configuring this script for automation.
Those variables are these if you need to find them:
    $CuckooIPandPort = ""
    $emailAddress = ""
    $folderChosen = ""
    $apikey = ""
#>

[CmdletBinding()]
[Alias()]
Param
(
    # Manual or Automated Option
    [Parameter(Mandatory=$false,
                ValueFromPipelineByPropertyName=$true,
                Position=0)]
    $automated,
    # Seconds to execute
    [Parameter(Mandatory=$false,
                ValueFromPipelineByPropertyName=$true,
                Position=1)]
    $duration
)

#Enter static configuration settings here for the automated version of this script!
function staticConf {
    #Automated variables to be declared here or from a file
    #CuckooIPandPort example: "http://192.168.0.20:8090"
    $CuckooIPandPort = ""
    #Email address used in Outlook example: "email@address.com"
    $emailAddress = ""
    #Folder name in your inbox example: "Cuckoo"
    $folderChosen = ""
    #API Key configured in your cuckoo.conf file for authentication, highly recommended that you use this if you are running Cuckoo 2.0.7+!
    $apiKey = ""
    if($CuckooIPandPort -eq "" -or $emailAddress -eq "" -or $folderChosen -eq ""){
        return $null
    }else{
        return $CuckooIPandPort, $emailAddress, $folderChosen, $apiKey
    }
}
switch ($automated) {
    $true {Write-Host "You have set automate to $automated! Good choice!"; $automated = $true}
    $false {Write-Host "You have set automate to $automated! Good choice!"; $automated = $false}
    $null {
        $automated = Read-Host -Prompt 'Automated version? (Enter true or false)'
        Write-Host "Option $automated"
        if($automated -ne $true -and $automated -ne $false){
            Write-Host "Not a valid option, please enter true or false. Exiting";Pause;exit
        }
    }
    default{Write-Host "Not a valid option, please enter true or false. Exiting";Pause;exit}
}

#Parse Email Message - Ready Outlook
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")

function checkCuckoo($CuckooIPandPort){
    try {
        if($apiKey){
            $headers = @{Authorization = "Bearer $apiKey"}
        }else{
            $headers = $null
        }
        $cuckooStatus = Invoke-RestMethod $CuckooIPandPort"/cuckoo/status" -Headers $headers
        Write-Host -ForegroundColor Green $cuckooStatus.hostname'(Version:'$cuckooStatus.version') loaded!'
    }
    catch {
        Write-Host -BackgroundColor Red "Could not get Cuckoo status...abandoning.."
        $errors = $_ | ConvertFrom-Json
        if($errors.message -contains "'Authorization: Bearer ' header is required"){Write-host "API Key is incorrect or not found! Double check your cuckoo.conf for the API key and verify it is the same key you entered for PowerCuckoo to use in the static configuration."}
        Exit
    }
}

switch ($automated){ 
    ($false){
        #Some things to load first - Pay no attention here
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic'); Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing; Add-Type -AssemblyName PresentationFramework
        #Cuckoo Config# 127.0.0.1:8090 is default, to change use: cuckoo api --host 127.0.0.1 -p 80

        $CuckooIPandPort = [Microsoft.VisualBasic.Interaction]::InputBox("Set your Cuckoo Host and Port`n`n`nFor Example: http://127.0.0.1:8090", "Where is your Cuckoo nesting?")
        $apiKey = [Microsoft.VisualBasic.Interaction]::InputBox("What is your API Key if you are using one?`n`n`nFor Example: 1BXU9tfuBgsY5R0KxrmDIg", "What is your Cuckoo's api key?")

        #Check to see if Cuckoo is configured and working appropriately
        checkCuckoo $CuckooIPandPort

        #Ask for Email Address for Outlook
        $title = 'Email Address Configuration'
        $msg = "Please Enter your Outlook Email Address`n`nExample: email@address.com`n`nTip: Make sure Outlook is Running!"
        $emailAddress = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

        #Ask for which folder to search
        $title = 'Email Folder Configuration'
        $foldersAvailable = $namespace.Folders.Item($emailAddress).Folders | Select-Object Name
        $msg = "Please Enter your Outlook Email Folder you wish to parse: $foldersAvailable"
        #Ask for folder input - also can be used to statically select folder name (see $foldername below)
        #All the GUI form data stuff\/ \/ \/
        $form = New-Object System.Windows.Forms.Form; $form.Text = 'Select a folder'
        $form.Size = New-Object System.Drawing.Size(300,200); $form.StartPosition = 'CenterScreen'
        $OKButton = New-Object System.Windows.Forms.Button; $OKButton.Location = New-Object System.Drawing.Point(75,120)
        $OKButton.Size = New-Object System.Drawing.Size(75,23); $OKButton.Text = 'OK'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.AcceptButton = $OKButton; $form.Controls.Add($OKButton)

        $CancelButton = New-Object System.Windows.Forms.Button; $CancelButton.Location = New-Object System.Drawing.Point(150,120)
        $CancelButton.Size = New-Object System.Drawing.Size(75,23); $CancelButton.Text = 'Cancel'
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.CancelButton = $CancelButton; $form.Controls.Add($CancelButton)

        $label = New-Object System.Windows.Forms.Label; $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(280,20); $label.Text = 'Please select a folder:'; $form.Controls.Add($label)

        $listBox = New-Object System.Windows.Forms.ListBox; $listBox.Location = New-Object System.Drawing.Point(10,40); 
        $listBox.Size = New-Object System.Drawing.Size(260,20); $listBox.Height = 80

        $foldersAvailable.Name | ForEach-Object {
            [void] $listBox.Items.Add("$_")
        }

        $form.Controls.Add($listBox); $form.Topmost = $true; $result = $form.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK){
                $folderChosen = $listBox.SelectedItem
                Write-Host "Folder Chosen: $folderChosen" -ForegroundColor Green
        }else{
                Write-Host "No folder Chosen!" -ForegroundColor Red
        }
        #All the GUI form data stuff/\ /\ /\
    }
    ($true){
        $staticVars = staticConf
        if($staticVars -eq $null){Write-Host "No static config found in PowerCuckoo.ps1 file. Adjust those variables. Going to exit."; pause; exit;
        }else{
            $staticVars = staticConf
            $CuckooIPandPort = $staticVars[0]
            $emailAddress = $staticVars[1]
            $folderChosen = $staticVars[2]
            $apiKey = $staticVars[3]
            $staticVars
            checkCuckoo $CuckooIPandPort
        }
    }

}

#Cuckoo REST API
$CuckooREST = $CuckooIPandPort+"/"
$MaliciousFileREST = $CuckooREST + 'tasks/create/file'
$MaliciousURLREST = $CuckooREST + 'tasks/create/url'
#$MaliciousArchiveREST = $CuckooREST + 'tasks/create/submit'

$folderName = $folderChosen
#RegEx to Grab URL
$RegExSpecial = '((http|ftp|https):\/\/([\w_-]+(?:(?:\.[\w_-]+)+))([\w.,@?^=%&:/~+#-;]*[\w@?^=%&/~+#-;])?)'
if($duration -eq $null){$runOnce = $true}
do{
#Cuckoo Folder - #Feed the Cuckoo Subfolder
$FeedTheCuckooUnread = $namespace.Folders.Item($emailAddress).Folders.Item($folderName).Items | Where-Object UnRead -EQ true
$unreadCount = $FeedTheCuckooUnread.Count
if(!$unreadCount){$unreadCount = $FeedTheCuckooUnread.UnRead.Count}
Write-Host -ForegroundColor Green "Found $unreadCount Unread Items to parse"

#Parse Email for URLs
function findURLs {
    $EmailBodyToSearch = @()
    $urlsFound = @()

    #Store URLs to get Searched from Email HTMLBody
    $EmailBodyToSearch += $FeedTheCuckooUnread.HTMLBody

    #Loop through results for URLs
    $EmailBodyToSearch | ForEach-Object {
    if ($_ -match $RegExSpecial)
        {
            $urlMatches = $_ | Select-String -AllMatches $RegExSpecial
            $urlsFound += $urlMatches.Matches.Value
        }
    }
    if($urlsFound.Count -ge 1){
        $totalURLs = $urlsFound.Count
        Write-Host "Found $totalURLs, now reducing to unique."
        $urlsFound = $urlsFound | Select-Object -Unique
        $uniqueTotalURLs = $urlsFound.Count
        $totalDiffURLs = $totalURLs-$uniqueTotalURLs
        Write-Host "Down to $uniqueTotalURLs found. Removed $totalDiffURLs duplicate(s)."
    }
    if(!([string]::IsNullOrEmpty($urlsFound))){
        if($automated -eq $false){
            [System.Windows.MessageBox]::Show('URL(s) Found! Going to Submit: ' + $uniqueTotalURLs + ' URLs to Cuckoo!!')
            $msgBoxInput =  [System.Windows.MessageBox]::Show('Would you like to see the URLs?','Urls Found','YesNo','Warning')

            switch  ($msgBoxInput) {
                'Yes' {
                    [System.Windows.MessageBox]::Show($urlsFound)
                    $x = 0
                    $urlsFound  | ForEach-Object {
                        if(!([string]::IsNullOrEmpty($_))){
                            Write-Host ($x+1)  $urlsFound[$x]
                            $x++
                        }
                    }
                }
                'No' {
                    [System.Windows.MessageBox]::Show('Okay, carry on then.')
                }
            }
        }else{Write-Host "URL(s) Found! Going to Submit: $uniqueTotalURLs URLs to Cuckoo!!"}

        maliciousURLSubmission $urlsFound

    }else{
        Write-Host 'No URLs found in Email with attachment.'
       
    }
}
 
#Send Cuckoo a malicious File
function maliciousFileSubmission ($submitFile) {
    $FilePath = Get-Location
    $FeedTheCuckooUnread | ForEach-Object{
        $messageSubject = $_.Subject
        Write-Host $messageSubject
        #Store each attachment then upload
        $_.Attachments | ForEach-Object {
            Write-Host "Attachment found: "$_.FileName -ForegroundColor Green
            $fileNameToStore = $_.FileName
            $tempFileStorage = (Join-Path $FilePath $fileNameToStore)
            $_.SaveAsFile($tempFileStorage)
            Write-Host "Attachment stored in: $tempFileStorage" -ForegroundColor Blue

            #Upload the file via REST API
            $fileBytes = [System.IO.File]::ReadAllBytes($tempFileStorage)
            $fileEnc = [System.Text.Encoding]::GetEncoding('UTF-8').GetString($fileBytes)
            $boundary = [System.Guid]::NewGuid().ToString()
            $LF = "`r`n";

            $bodyLines = ( 
                "--$boundary",
                "Content-Disposition: form-data; name=`"file`"; filename=`"$fileNameToStore`"",
                "Content-Type: application/octet-stream$LF",
                $fileEnc,
                "--$boundary--$LF" 
            ) -join $LF
            $task = ''
            #Send the encoded blob to Cuckoo!
            $task = Invoke-RestMethod -Uri $MaliciousFileREST -Method Post -ContentType "multipart/form-data; boundary=`"$boundary`"" -Body $bodyLines -Headers $headers
            if($automated -eq $false){[System.Windows.MessageBox]::Show('Running Malicious File Submission')}

            #If task submits successfully, delete the temporary created file.
            if($task){
                Write-Host -ForegroundColor Red "Deleting temporary file download...$fileNameToStore"
                Remove-Item $fileNameToStore
                if($automated -eq $false){
                [System.Windows.MessageBox]::Show("File analysis submitted for the Email Subject:`n"+$FeedTheCuckooUnread.Subject+"`n`nThe TaskID is: "+$task.task_id)
                }else{Write-Host "File analysis submitted for the Email Subject:"$FeedTheCuckooUnread.Subject"`n`nThe TaskID is:"$task.task_id}
            }else{
                Write-Host -ForegroundColor Yellow "There was an issue trying to submit a file to Cuckoo, it was not removed."
            }
        }
    }
}
 
#Function for sending Cuckoo malicious URLs
function maliciousURLSubmission ($submitURL) {
    if($automated -eq $false){
        [System.Windows.MessageBox]::Show('Running Malicious URL Submission')
    }else{Write-Host "Running Malicious URL Submission"}

    #Loop through all the URLs
    $submitURL | ForEach-Object {
        $submitURLx = $_
        $task = Invoke-RestMethod -Method Post -Uri $MaliciousURLREST -Body url=$submitURLx -Headers $headers
        $taskID = $task.task_id
        Write-Host "Task ID: $taskID `nURL Submitted: $submitURLx"
    }
    if($automated -eq $false){
    [System.Windows.MessageBox]::Show("All URLs ($uniqueTotalURLs) have been sent to Cuckoo!")
    }else{Write-Host "All URLs ($uniqueTotalURLs) have been sent to Cuckoo!"}

}

#Check for Attachments / URLs
$attachmentCount = $FeedTheCuckooUnread.Attachments.Count
#Mark Message(s) as Read 
$FeedTheCuckooUnread | ForEach-Object {
    $_.Unread = "False"
}
if ($attachmentCount -ge 1) {
    Write-Host "Attachments Found: $attachmentCount"
    maliciousFileSubmission
    Write-Host "On to looking for URLs"
    findURLs
}elseif($attachmentCount -eq 0){
    if($automated -eq $false){
        [System.Windows.MessageBox]::Show('No attachments, finding URLs for analysis')
    }else{Write-Host "No attachments, finding URLs for analysis"}
    #Find URL in Email
    findURLs
}else{
    if($automated -eq $false){
        [System.Windows.MessageBox]::Show('Something went terribly wrong')
    }else{Write-Host "Something went terribly wrong"}
}
Start-Sleep -Seconds $duration
$runOnce = $false
}while($duration -ge 0 -or $runOnce -eq $true)
Read-host "PowerCuckoo has finished running! Hit Enter to exit."
