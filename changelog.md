# Changelog
## 2.0.2 - 05/05/2020
 - Updating module to connect using "Basic" auth to on-premises (and discoverying endpoint by Autodiscover) and connect using "OAUTH" to Exchange Online.
 - Adding EWS API DLL to module folder, so no need to pre install EWS API DLL.
 - Minor fix on Process17 to clear HTTPHeader
## 2.0.0 - 05/04/2020
 - Moving tool as a Module in GitHub
 - Added ADAL capabilities as many organizations have MFA or even disabling Basic Auth.
## 1.8.2 - 02/15/2019
 - Added 2 columns to folder lists methods 1-6 : TotalItemsCount, # of Subfolders
## 1.8.0 - 04/16/2018
 - Optimized logon options. If we choose 'Office 365', we will not use SCP and we hard-code EXO endpoint.
## 1.7.0 - 03/13/2018
 - Added new method "remove OWA Configurations"
 - Added $service.HttpHeaders.Add("X-AnchorMailbox", $TargetSmtpAddress) , for better Impersonation performance
 - Removed AutodiscoverSCPLookup, as most of the uses of this app is with Exchange Online
## 1.0.0 - 02/19/2018
 - First Release
## 1.0.0 - 01/30/2018
 - Project start