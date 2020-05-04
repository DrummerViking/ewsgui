# Changelog
## 1.00 - 01/30/2018 - Project start
## 1.00 - 02/19/2018 - First Release
## 1.70 - 03/13/2018 - Added new method "remove OWA Configurations"
 - Added $service.HttpHeaders.Add("X-AnchorMailbox", $TargetSmtpAddress) , for better Impersonation performance
 - Removed AutodiscoverSCPLookup, as most of the uses of this app is with Exchange Online
## 1.80 - 04/16/2018 - Optimized logon options. If we choose 'Office 365', we will not use SCP and we hard-code EXO endpoint.
## 1.82 - 02/15/2019 - Added 2 columns to folder lists methods 1-6 : TotalItemsCount, # of Subfolders
## 2.00 - 04/29/2020 - Moving tool as a Module in GitHub
 - Added ADAL capabilities as many organizations have MFA or even disabling Basic Auth.