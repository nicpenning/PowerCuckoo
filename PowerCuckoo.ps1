<#PowerCuckoo
    Created by Nicholas Penning
    Date: 8/14/2017
    Updated: 9/13/2017
    Description: For automation!

    Note: Works well grabbing a folder from an email. Try it out!
#>
 
#Cuckoo REST Calls
$CuckooREST = 'http://localhost:8090'
$MaliciousFileREST = $CuckooREST + 'tasks/create/file'
$MaliciousURLREST = $CuckooREST + 'tasks/create/url'
$MaliciousArchiveREST = $CuckooREST + 'tasks/create/submit'
 
#Malzoo API Calls
#curl -X POST -F file=@/path/to/sample -F tag=yourtaghere http://localhost:1338/file/add

#Parse Email Message - Ready Outlook
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")
#Ask for Email Address for Outlook
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$title = 'Email Address Configuration'
$msg = 'Enter your Outlook Email Address'
$emailAddress = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

#Ask for which folder to search
$title = 'Email Folder Configuration'
$msg = 'Enter your Outlook Email Folder you wish to parse'
$folderName = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
 
#RegEx to Grab URL
$RegExHtmlLinks = '<a\s+(?:[^>]*?\s+)?href="[h+f]([^"]*)"'
$RegExHostName = '^(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])$'
 
#Cuckoo Folder - #Feed the Cuckoo Subfolder
$FeedTheCuckooUnread = $namespace.Folders.Item($emailAddress).Folders.Item($folderName).Items | Where-Object UnRead -EQ true
 
#Save Email as .MSG file
function saveEmail{
    #The following command will save only the unread emails in the Inbox to the C:\Saved Emails\ folder
    $emailDestinationPath = 'C:\Saved Emails\Inbox'
    $outbox = 'C:\Saved Emails\Outbox'
    .\Save-EmailCuckoo.ps1 -DestinationPath $emailDestinationPath -UnreadOnly #-MarkRead
    $emailMessage = Get-ChildItem $emailDestinationPath
    $emailPath = $emailDestinationPath+'\'+$emailMessage.Name
 
    #Submit for Analysis
    maliciousFileSubmission($emailPath)
}
 
#Parse Email for URLs
function findURLs {
    $EmailBodyToSearch = @()
    $urlsFound = @()
   
    #Mark Message as Read 
    $x = 0
    $FeedTheCuckooUnread.UnRead | ForEach-Object {
        $FeedTheCuckooUnread[$x].UnRead = "False"
        $x++
    }

    #Store URLs to get Searched from Email HTMLBody
    $EmailBodyToSearch += $FeedTheCuckooUnread.HTMLBody

    #Loop through results for URLs
    $EmailBodyToSearch | ForEach-Object {
    if ($_ -match $RegExHtmlLinks)
        {
            $urlsFound += $matches[0]
        }
    }
    $urlsFound = $urlsFound | select-string -pattern $RegExHtmlLinks -AllMatches | Foreach {$_.Matches} | ForEach-Object {$_.Value} | Select-Object -Unique
 
    #Clean URLs for Analysis
    $cleanedUrlsForAnalysis = $urlsFound -replace '<a href='
    $cleanedUrlsForAnalysis = $cleanedUrlsForAnalysis -replace '"'

    if(!([string]::IsNullOrEmpty($cleanedUrlsForAnalysis))){
        #Write-Host 'URL Found! Going to Submit:' $cleanedUrlsForAnalysis.Count'URLs to Cuckoo!!'
        [System.Windows.MessageBox]::Show('URL(s) Found! Going to Submit: ' + $cleanedUrlsForAnalysis.Count + ' URLs to Cuckoo!!')
        $msgBoxInput =  [System.Windows.MessageBox]::Show('Would you like to see the URLs?','Urls Found','YesNo','Warning')

        switch  ($msgBoxInput) {
            'Yes' {
                [System.Windows.MessageBox]::Show($cleanedUrlsForAnalysis)
                $x = 0
                $cleanedUrlsForAnalysis | ForEach-Object {
                if(!([string]::IsNullOrEmpty($_))){
                    Write-Host ($x+1) $cleanedUrlsForAnalysis[$x]
                $x++
            }
        }
            }
            'No' {
                [System.Windows.MessageBox]::Show('Okay, carry on then.')
            }
        }

        maliciousURLSubmission ($cleanedUrlsForAnalysis)

    }else{
        Write-Host 'No URLs found in Email with attachment.'
       
    }
}
 
#Send Cuckoo a malicious File
function maliciousFileSubmission ($submitFile) {
    [System.Windows.MessageBox]::Show('Running Malicious File Submission')
    #$analysis = .\curl.exe -F file=@$submitFile $MaliciousFileREST
    $analysis = $analysis -replace '{'
    $analysis = $analysis -replace '}'
    $analysis = $analysis -replace ' '
    $analysis = $analysis -replace '"'
   [System.Windows.MessageBox]::Show('File analysis submitted: '+ $emailMessage.Name + ' with '+ $analysis)
    #move $emailPath $outbox
}
 
#Function for sending Cuckoo malicious URLs
function maliciousURLSubmission ($submitURL) {
    [System.Windows.MessageBox]::Show('Running Malicious URL Submission')
    #Invoke-RestMethod -Method Post -Uri $MaliciousURLREST -Body url=$MaliciousSite
    $x = 0
    #Loop through all the URLs in the cleaned up array
    $submitURL | ForEach-Object {
        $submitURLx = $submitURL[$x]
        #Invoke-RestMethod -Method Post -Uri $MaliciousURLREST -Body url=$submitURLx
        [System.Windows.MessageBox]::Show($submitURLx)
        $x++
    }
}
 
#Send Email to CuckooFeeder
function sendEmail{
    #Send Email Setup
    #$Mail = $Outlook.CreateItem(0)
    #$Mail.To =
    #$Mail.Subject = $FeedTheCuckooUnread.Subject
    #$Mail.HTMLBody = $FeedTheCuckooUnread.HTMLBody
    #$Mail.Send()
    Write-Host 'Sending Email'
    #Cleanup Mail Variable
    Remove-Variable Mail
}
 
#Check for Attachments / URLs
if ($FeedTheCuckooUnread.Attachments.Count -ge 1) {
    #Submit .MSG file to CuckooFeeder
    sendEmail
    [System.Windows.MessageBox]::Show('Email Sent: ' + $FeedTheCuckooUnread.Subject + ' with ' + $FeedTheCuckooUnread.Attachments.Count + 'attachment(s). Now searching for URLs')
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
##ReportSpam
#$namespace.Folders.Item(2).Folders.Item(1).Folders.Item(1).Items | Where-Object UnRead -EQ True
# | Format-List Unread, CreationTime, SenderName, ConversationTopic, Body, HTMLBody, To
#Sample Data
#$MaliciousSite = "http://google.com"
#$MaliciousFile = ".\Alert.msg"
#>
