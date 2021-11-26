# Changelog
## 2.0.13 - 11/26/2021
 - Updated authentication libraries from ADAL to MSAL.
## 2.0.12 - 11/16/2021
 - Adding requirement to run the module in PS 'Desktop' Edition.
## 2.0.11 - 09/17/2021
 - Adding auto-check for update functionality.
## 2.0.10 - 05/21/2020
 - Updating project's build scripts.  
## 2.0.9 - 05/21/2020
 - Minor change in ExchangeVersion number for EWS Object.
 - Removing AzureAD Module dependency on this module installation.  
## 2.0.8 - 05/07/2020
 - Moved all methods to individual Functions.
 - Removed unused objects from Start-EWSGUi function.  
## 2.0.6 - 05/06/2020
 - Updating Readme files.
 - Removing old function "Import-EWSDLL". Not needed anymore.
 - Minor changes on PSFMessages after each method.
 - Moving Method1to6 to individual function.  
## 2.0.2 - 05/05/2020
 - Updating module to connect using "Basic" auth to on-premises (and discoverying endpoint by Autodiscover) and connect using "OAUTH" to Exchange Online.
 - Adding EWS API DLL to module folder, so no need to pre install EWS API DLL.
 - Minor fix on Method17 to clear HTTPHeaders.  
## 2.0.0 - 05/04/2020
 - Moving tool as a Module in GitHub
 - Added ADAL capabilities as many organizations have MFA or even disabling Basic Auth.  
## 1.8.2 - 02/15/2019
 - Added 2 columns to folder lists methods 1-6 : TotalItemsCount, # of Subfolders.  
## 1.8.0 - 04/16/2018
 - Optimized logon options. If we choose 'Office 365', we will not use SCP and we hard-code EXO endpoint.  
## 1.7.0 - 03/13/2018
 - Added new method "remove OWA Configurations".
 - Added $service.HttpHeaders.Add("X-AnchorMailbox", $TargetSmtpAddress) , for better Impersonation performance.
 - Removed AutodiscoverSCPLookup, as most of the uses of this app is with Exchange Online.  
## 1.0.0 - 02/19/2018
 - First Release.  
## 1.0.0 - 01/30/2018
 - Project start.