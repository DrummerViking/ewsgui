# EWS-GUI Tool

## About
Exchange Web Services (EWS) tool to perform different operations in Exchange On-Premises and Exchange Online.  
This tool will connect using Basic Auth to on-premises mailboxes (uses Autodiscover for endpoint discovery).  
And will use Oauth to connect to Exchange Online. If "Modern Authentication" is not enabled in the tenant, the tool will fail to connect.  

## Pre-requisites

 > This Module requires Powershell 5.1 and above. It should work fine in PS7 and PS5.1.  
 > This Module will install Microsoft.Identity.Client module, in order to use MSAL libraries to connect to Exchange Online.  
 > There is no need to have EWS API Management dll pre installed. This Module already has the required files.  

## Installation

Opening Powershell with "Run as Administrator" you can run:
``` powershell
Install-Module EWSGui -Force
```
Once the module is installed, you can run:
``` powershell
Start-EWSGui
```

If you want to check for module updates you can run (the tool will already check for updates automatically):
``` powershell
Find-Module EWSGui
```
If there is any newer version than the one you already have, you can run:
``` powershell
Update-Module EWSGui -Force
```

## Module features:
### Allows to perform 16 different operations using EWS API:
- Option 1 : List Folders in Root
- Option 2 : List Folders in Archive Root
- Option 3 : List Folders in Public Folder Root
- Option 4 : List subFolders from a desired Parent Folder
- Option 5 : List folders in Recoverable Items Root folder
- Option 6 : List folders in Recoverable Items folder in Archive
- Option 7 : List Items in a desired Folder
- Option 8 : Create a custom Folder in Root
- Option 9 : Delete a Folder
- Option 10 : Get user's Inbox Rules
- Option 11 : Get user's OOF Settings
- Option 12 : Move items between folders
- Option 13 : Delete a subset of items in a folder
- Option 14 : Get user's Delegate information
- Option 15 : Change sensitivity to items in a folder
- Option 16 : Remove OWA configurations
- Option 17 : Switch to another Mailbox

## Version History
[Change Log](/ewsgui/changelog.md)