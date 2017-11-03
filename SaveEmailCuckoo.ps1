<#
.SYNOPSIS
This script will save emails in the .msg format from the Outlook Inbox and place them in a folder of your choosing.
 
.DESCRIPTION
Using this script will save emails to the path specified in the DestinationFolder parameter.  If you use the UnreadOnly switch, only unread emails will be saved.  Otherwise, all emails will be saved. The MarkRead switch will mark unread email as read, but can only be used with the UnreadOnly parameter.  The emails will be saved as .msg files and can be opened with Outlook.
 
.PARAMETER DestinationPath
Specifies the path to the destination folder where the emails will be saved
 
.PARAMETER UnreadOnly
If used, only unread emails will be saved.  If not, all emails will be saved.
 
.PARAMETER MarkRead
If used, will mark unread emails as read.  Can only be used with the UnreadOnly parameter.
 
.INPUTS
System.IO.DirectoryInfo.  Will bind the property FullName from a directory object to the DestinationPath parameter in the pipeline.
 
.OUTPUTS
None.  Save-Email.ps1 does not generate any output.
 
.EXAMPLE
The following command will save all the emails in the Inbox to the C:\Saved Emails\ folder
 
PS C:\> .\Save-Email.ps1 -DestinationPath 'C:\Saved Emails'
 
.EXAMPLE
The following command will save only the unread emails in the Inbox to the C:\Saved Emails\ folder
 
PS C:\> .\Save-Email.ps1 -DestinationPath 'C:\Saved Emails' -UnreadOnly
 
.EXAMPLE
The following command will save only the unread emails in the Inbox to the C:\Saved Emails\ folder and mark them as read
 
PS C:\> .\Save-Email.ps1 -DestinationPath 'C:\Saved Emails' -UnreadOnly -MarkRead
 
.EXAMPLE
This example demonstrates using the pipeline to send a DirectoryInfo object to the script in the pipeline
 
PS C:\> Get-ChildItem 'C:\Users' -Recurse | Where-Object {$_.Name -eq "Saved Emails"} | .\Save-Email.ps1
 
.NOTES
If you do not enter a value for the DestinationPath parameter on the command line, you will be prompted to enter a value after pressing <Enter>.  At this point, do not enclose the value with quotation marks even if the path includes spaces. 
 
.LINK
http://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.application.aspx
http://blogs.technet.com/b/heyscriptingguy/archive/2011/05/26/use-powershell-to-data-mine-your-outlook-inbox.aspx
#>
 
#Binding for Common Parameters
[CmdletBinding(DefaultParameterSetName="All")]
 
Param(
    [Parameter(Mandatory=$true,
        Position=0,
        HelpMessage='Folder path to store emails. Do not use quotation marks even if the path has spaces.',
        ValueFromPipelineByPropertyName=$true
    )]
    [Alias("Destination", "Dest", "FullName")]
    [String]$DestinationPath,
     
    [Parameter(ParameterSetName="All")]
    [Parameter(Mandatory=$true,ParameterSetName="Unread")]
    [Switch]$UnreadOnly,
 
    [Parameter(ParameterSetName="Unread")]
    [Switch]$MarkRead
)
 
#Removes invalid Characters for file names from a string input and outputs the clean string
#Similar to VBA CleanString() Method
#Currently set to replace all illegal characters with a hyphen (-)
Function Remove-InvalidFileNameChars {
 
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [String]$Name
    )
 
    return [RegEx]::Replace($Name, "[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars())), '-')
}
 
#Test for destination folder nonexistence
if (!(Test-Path $DestinationPath)) {
    #Set values for prompt and menu
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
        "Confirmation Choice"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
        "Negative Response"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    $title = "Invalid Destination"
    $message = "The folder you entered does not exist.  Would you like to create the folder?"
 
    #Prompt for folder creation and store answer
    $result = $host.UI.PromptForChoice($title, $message, $options, 0)
 
    #If yes, create.
    if ($result -eq 0) {
        New-Item $DestinationPath -ItemType Directory | Out-Null
        Write-Host "Directory created."
    }
    #If no, exit
    else {exit}
}
     
#Add a trailing "\" to the destination path if it doesn't already
if ($DestinationPath[-1] -ne "\") {
    $DestinationPath += "\"
}
 
#Add Interop Assembly
Add-type -AssemblyName "Microsoft.Office.Interop.Outlook" | Out-Null
 
#Type declaration for Outlook Enumerations, Thank you Hey, Scripting Guy! blog for this demonstration
$olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
$olSaveType = "Microsoft.Office.Interop.Outlook.OlSaveAsType" -as [type]
$olClass = "Microsoft.Office.Interop.Outlook.OlObjectClass" -as [type]
 
#Add Outlook Com Object, MAPI namespace, and set folder to the Inbox
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNameSpace("MAPI")
#Folders specified from PowerCuckoo - Cuckoo - Feed the Cuckoo
$folder = $namespace.Folders.Item($emailAddress).Folders.Item($folderName).Items
 
#Iterate through each object in the chosen folder
foreach ($email in $folder.Items) {
     
    #Get email's subject and date
    [string]$subject = $email.Subject
    [string]$sentOn = $email.SentOn
    #Clean out '[]' and ','
    $subjectClean = $email.Subject -replace '[[\]]'
    $subjectCleaned = $subjectClean -replace ','
    #Strip subject and date of illegal characters, add .msg extension, and combine
    $fileName = Remove-InvalidFileNameChars -Name ($sentOn + "-" + $subjectCleaned + ".msg")
 
    #Combine destination path with stripped file name
    $dest = $DestinationPath + $fileName
     
    #Test if object is a MailItem
    if ($email.Class -eq $olClass::olMail) {
         
        #Test if UnreadOnly switch was used
        if ($UnreadOnly) {
             
            #Test if email is unread and save if true
            if ($email.Unread) {
                 
                #Test if MarkRead switch was used and mark read
                if ($MarkRead) {
                    $email.Unread = $false
                }
                $email.SaveAs($dest, $olSaveType::olMSG)
            }
        }
        #UnreadOnly switch not used, save all
        else {
            $email.SaveAs($dest, $olSaveType::olMSG)
        }
    }
}
 
#$outlook.Quit()
Remove-Variable folder
Remove-Variable namespace
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
Remove-Variable outlook
