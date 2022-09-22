# EWS-GUI Tool

## About
Exchange Web Services (EWS) tool to perform different operations in Exchange On-Premises and Exchange Online.  
This tool will connect using Basic Auth to on-premises mailboxes (uses Autodiscover for endpoint discovery).  
And will use Oauth to connect to Exchange Online. If "Modern Authentication" is not enabled in the tenant, the tool will fail to connect.  

## Pre-requisites

 > This Module requires Powershell 5.1 and above. It should work fine in PS7 and PS5.1.  
 > This Module will install 'Microsoft.Identity.Client' and 'MSAL.PS' modules, in order to use MSAL libraries to connect to Exchange Online.  
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

## Authentication options

The EWSGui tool can connect to Exchange On-premises mailboxes using Basic Authentication.  
To connect to Exchange Online, it will use Modern auth and we have 2 options, either with Delegated Permission or Application permission.  
Please check on the following page for more details and options to configure your EWSGui module.
[Authentication Options](/docs/AuthenticationOptions.md)  

## Module features:
### Allows to perform 16 different operations using EWS API:
- Option 1 : List Folders in Root
- Option 2 : List Folders in Archive Root
- Option 3 : List Folders in Public Folder Root
- Option 4 : List folders in Recoverable Items Root folder
- Option 5 : List folders in Recoverable Items folder in Archive
- Option 6 : List Items in a desired Folder
- Option 7 : Create a custom Folder in Root
- Option 8 : Delete a Folder
- Option 9 : Get user's Inbox Rules
- Option 10 : Get user's OOF Settings
- Option 11 : Move items between folders
- Option 12 : Delete a subset of items in a folder
- Option 13 : Get user's Delegate information
- Option 14 : Change sensitivity to items in a folder
- Option 15 : Remove OWA configurations
- Option 16 : Switch to another Mailbox

## Module logging

The module offers the command `Export-EWSGuiLog` in order to export module logs to CSV file and/or to Powershell GridView.  
More info [here](/docs/Export-EWSGuiLog.md).  

## Version History
[Change Log](/ewsgui/changelog.md)