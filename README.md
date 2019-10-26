
<p align="center">
  <img src="https://github.com/nicpenning/PowerCuckoo/blob/master/images/PowerCuckooLogo.png">
</p>

***
# PowerCuckoo
Using PowerShell on Windows with Outlook to interact with Cuckoo

What's this?
- Customizable script that reads Outlook email using the Outlook ComObject to read, parse, and send emails to Cuckoo for analysis.

Requirements:
 - Windows 7+
 - Outlook 2010+ (Running)
 - PowerShell
 - Cuckoo 2.0+ (API running)

Setup
 - Download PowerCuckoo.ps1
 - Open Outlook
 - Run ./PowerCuckoo.ps1

What works:
 - Tested on Windows 10 with Office 365 (Version 1901)
 - Reading a specific email folder to look for URLs or Attachments and submits them!
 - Warning: Becareful using the built-in Junk Email folder, for it may distort URLs/Attachments for analysis!

Usage:
 - ./PowerCuckoo.ps1
 - Check out the Wiki!
 https://github.com/nicpenning/PowerCuckoo/wiki
 
TODO:
  - Add Auto-Setup/Install
  - Create EWS version
  - Add SSL/TLS Support
  - Add Alerting/Reporting

If you would like to join our Teams channel please send a request to this address:
7d98f163.Yamsec401.onmicrosoft.com@amer.teams.ms
