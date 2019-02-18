<#
.SYNOPSIS
Send URLs from an Outlook Email folder that contains unread messages to Cuckoo.

.DESCRIPTION
    PowerCuckoo
    Created by Nicholas Penning
    Date: 8/14/2017
    Updated: 2/18/2019

    This script is currently gui/manually driven but can be automated by statically setting some variables.
    This initial release is for testing to get an understanding of how it works. 
    The goal is to create a fully automated version that is liteweight and easy to use.

.EXAMPLE
./PowerCuckoo.ps1

.NOTES
Currently works well for grabbing Unread messages from a folder of your choosing and sending them to your Cuckoo host. 
Try it out!
#>
#Some things to load first - Pay no attention here
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
#Cuckoo config
#127.0.0.1:8090 is default, to change use: cuckoo api --host 127.0.0.1 -p 80
try {
    $CuckooIPandPort = [Microsoft.VisualBasic.Interaction]::InputBox("Set your Cuckoo Host and Port`n`n`nFor Example: http://127.0.0.1:8090", "Where is your Cuckoo nesting?")
    $cuckooStatus = Invoke-RestMethod $CuckooIPandPort"/cuckoo/status"
    Write-Host -ForegroundColor Green $cuckooStatus.hostname'(Version:'$cuckooStatus.version') loaded!'
}
catch {
    Write-Host -BackgroundColor Red "Could not get Cuckoo status...abandoning.."
    Exit
}

#Cuckoo REST API
$CuckooREST = $CuckooIPandPort+"/"
#$MaliciousFileREST = $CuckooREST + 'tasks/create/file'
$MaliciousURLREST = $CuckooREST + 'tasks/create/url'
#$MaliciousArchiveREST = $CuckooREST + 'tasks/create/submit'
 
#Malzoo API Calls
#curl -X POST -F file=@/path/to/sample -F tag=yourtaghere http://localhost:1338/file/add

#Parse Email Message - Ready Outlook
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")
#Ask for Email Address for Outlook
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$title = 'Email Address Configuration'
$msg = "Enter your Outlook Email Address`n`nTip: Make sure Outlook is Running!"
$emailAddress = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

#Ask for which folder to search
$title = 'Email Folder Configuration'
$foldersAvailable = $namespace.Folders.Item($emailAddress).Folders | Select-Object Name
$msg = "Enter your Outlook Email Folder you wish to parse: $foldersAvailable"
#Manually ask for folder input - also can be used to statically select folder name
#$folderName = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select a folder'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please select a folder:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

$foldersAvailable.Name | ForEach-Object {
    [void] $listBox.Items.Add("$_")
}

$form.Controls.Add($listBox)
$form.Topmost = $true
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK){
        $folderChoosen = $listBox.SelectedItem
        Write-Host "Folder choosen: $folderChoosen" -ForegroundColor Green
}else{
        Write-Host "No folder choosen!" -ForegroundColor Red
}
$folderName = $folderChoosen
#RegEx to Grab URL
#$RegExHtmlLinks = '<a\s+(?:[^>]*?\s+)?href="([h+f][^"]*)"'
#$RegExHostName = '^(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])$'
$RegExSpecial = '((http|ftp|https):\/\/([\w_-]+(?:(?:\.[\w_-]+)+))([\w.,@?^=%&:/~+#-;]*[\w@?^=%&/~+#-;])?)'
#Cuckoo Folder - #Feed the Cuckoo Subfolder
$FeedTheCuckooUnread = $namespace.Folders.Item($emailAddress).Folders.Item($folderName).Items | Where-Object UnRead -EQ true
$unreadCount = $FeedTheCuckooUnread.Count
Write-Host -ForegroundColor Green "Found $unreadCount Unread Items to parse"

#Save Email as .MSG file
<#function saveEmail{
    #The following command will save only the unread emails in the Inbox to the C:\Saved Emails\ folder
    $emailDestinationPath = 'C:\Saved Emails\Inbox'
    #$outbox = 'C:\Saved Emails\Outbox'
    .\Save-EmailCuckoo.ps1 -DestinationPath $emailDestinationPath -UnreadOnly #-MarkRead
    $emailMessage = Get-ChildItem $emailDestinationPath
    $emailPath = $emailDestinationPath+'\'+$emailMessage.Name
 
    #Submit for Analysis
    maliciousFileSubmission($emailPath)
}#>

#Parse Email for URLs
function findURLs {
    $EmailBodyToSearch = @()
    $urlsFound = @()
   
    #Mark Message(s) as Read 
    $FeedTheCuckooUnread | ForEach-Object {
        $_.Unread = "False"
    }

    #Store URLs to get Searched from Email HTMLBody
    $EmailBodyToSearch += $FeedTheCuckooUnread.HTMLBody
    #$testURLFind = $EmailBodyToSearch | Select-String -AllMatches $RegExSpecial
    #Loop through results for URLs
    $EmailBodyToSearch | ForEach-Object {
    if ($_ -match $RegExSpecial)
        {
            $urlMatches = $_ | Select-String -AllMatches $RegExSpecial
            $urlsFound += $urlMatches.Matches.Value
        }
    }
    $totalURLs = $urlsFound.Count
    Write-Host "Found $totalURLs, now reducing to unique."
    $urlsFound = $urlsFound | Select-Object -Unique
    $uniqueTotalURLs = $urlsFound.Count
    $totalDiffURLs = $totalURLs-$uniqueTotalURLs
    Write-Host "Down to $uniqueTotalURLs found. Removed $totalDiffURLs duplicate(s)."
    #Clean URLs for Analysis

    if(!([string]::IsNullOrEmpty($urlsFound))){
        #Write-Host 'URL Found! Going to Submit:' $cleanedUrlsForAnalysis.Count'URLs to Cuckoo!!'
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

        maliciousURLSubmission $urlsFound

    }else{
        Write-Host 'No URLs found in Email with attachment.'
       
    }
}
 
#Send Cuckoo a malicious File
<#function maliciousFileSubmission ($submitFile) {
    [System.Windows.MessageBox]::Show('Running Malicious File Submission')
    #$analysis = .\curl.exe -F file=@$submitFile $MaliciousFileREST
    $analysis = $analysis -replace '{'
    $analysis = $analysis -replace '}'
    $analysis = $analysis -replace ' '
    $analysis = $analysis -replace '"'
   [System.Windows.MessageBox]::Show('File analysis submitted: '+ $emailMessage.Name + ' with '+ $analysis)
    #move $emailPath $outbox
}#>
 
#Function for sending Cuckoo malicious URLs
function maliciousURLSubmission ($submitURL) {
    [System.Windows.MessageBox]::Show('Running Malicious URL Submission')

    #Loop through all the URLs
    $submitURL | ForEach-Object {
        $submitURLx = $_
        $task = Invoke-RestMethod -Method Post -Uri $MaliciousURLREST -Body url=$submitURLx
        $taskID = $task.task_id
        [System.Windows.MessageBox]::Show("Task ID: $taskID"+"`nURL Submitted: $submitURLx")
    }

    [System.Windows.MessageBox]::Show("All URLs ($uniqueTotalURLs) have been sent to Cuckoo!")

}
 
#Check for Attachments / URLs
$attachmentCount = $FeedTheCuckooUnread.Attachments.Count

if ($FeedTheCuckooUnread.Attachments.Count -ge 1) {
    Write-Host "Attachments Found: $attachmentCount (Not currently supported to send for analysis)"
    Write-Host "Will look for URLs now."
    findURLs
}elseif($FeedTheCuckooUnread.Attachments.Count -eq 0){
   [System.Windows.MessageBox]::Show('No attachments, finding URLs for analysis')
    #Find URL in Email
    findURLs
}else{
    [System.Windows.MessageBox]::Show('Something went terribly wrong')
}
 
<#
Other Tests and Useful Commands
#$namespace.Folders.Item(2).Folders.Item(1).Folders.Item(1).Items | Where-Object UnRead -EQ True
# | Format-List Unread, CreationTime, SenderName, ConversationTopic, Body, HTMLBody, To
#Sample Data
#$MaliciousSite = "http://google.com"
#$MaliciousFile = ".\Alert.msg"
#>
