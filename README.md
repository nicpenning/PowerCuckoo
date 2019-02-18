# PowerCuckoo
Using PowerShell to interact with Cuckoo

What's this?
- Customizable script that reads Outlook email using the Outlook ComObject to read, parse, and send emails to Cuckoo for analysis

Much more to come!

Requirements:
 - Windows 7+
 - Outlook 2010+
 - PowerShell
 - SaveEmailCuckoo.ps1
 - curl.exe

Setup

 - Create a PowerCuckoo directory and store PowerCuckoo.ps1, SaveEmailCuckoo.ps1, and curl.exe in that direcory.

Note: Not fully functioning.

What works:
 - Reading a specific email folder to look for URLs or Attachments

Usage:
 - ./PowerCuckoo.ps1
 
 TODO:
  - Update this README
  - Add screenshots of POC
  - Remove Curl.exe requirement
  - Remove SaveEmailCuckoo.ps1 requirement
  - Add Auto-setup/install
