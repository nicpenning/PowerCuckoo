# PowerCuckoo
Using PowerShell on Windows with Outlook to interact with Cuckoo

What's this?
- Customizable script that reads Outlook email using the Outlook ComObject to read, parse, and send emails to Cuckoo for analysis

Much more to come!

Requirements:
 - Windows 7+
 - Outlook 2010+ (Running)
 - PowerShell
 - Cuckoo 2.0+ (API running)

Setup
 - Download PowerCuckoo.ps1
 - Run ./PowerCuckoo.ps1

What works:
 - Tested on Windows 10 with Office 365 (Version 1901)
 - Reading a specific email folder to look for URLs or Attachments and submits them!

Usage:
 - ./PowerCuckoo.ps1
 - Check out the Wiki!
 https://github.com/nicpenning/PowerCuckoo/wiki
 
 TODO:
  - Add Auto-setup/install
  - Create Automated version
  - Create EWS version
