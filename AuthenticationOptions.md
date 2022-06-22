# Authentication options

## Using Delegated Permissions  

In order to connect using Modern Authentication with Delegated permissions, we need to have an Azure App registered with the assign permission.  
We are currently using a multi-tenant app with the permission to work with EWS. So you don't need to register any additional app in Azure.  
Once you run the app for the first time, you will have a pop-up window to consent App request.  
The app will request the user credentials and logon.  

If you are going to work with your own mailbox, nothing else is needed.   
If you want to switch and impersonate additional mailboxes you need to grant this impersonation permission to a user account first in EXO powershell by running:  
```Powershell
New-ManagementRoleAssignment -Name "Impersonation assignment name" -Role ApplicationImpersonation -User "account@domain.com"
```

## Using Application Permissions

In order to connect with an Application Permission, no ApplicationImpersonation role is needed in EXO.  
Just need to create the Application in Azure following: [Authenticate an EWS application by using OAuth](https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)  

Once you create your app with a ClientSecret, you can use this tool by running:  
```Powershell
Start-EWSGui -ClientID "your app client ID" -TenantID "Your tenant ID" -ClientSecret "your Secret passcode"
```

## Saving your Azure App details in the EWSGui module

If you want to use Application permission flow, we have an option to save your "ClientID", "TenantID" and "ClientSecret", so you don't need to enter it every time as the example above.  
you can run:  
```Powershell
Register-EWSGuiTenantData -ClientID "your app client ID" -TenantID "Your tenant ID" -ClientSecret "your Secret passcode"
```

Now everytime you want to run the module, just run `Start-EWSGui` and will fetch these saved details (so it will follow the Application permissions flow).  
if you need to revert this change, let's say you need to try Delegated Permission back again, you can unregister these values:  
```Powershell
Unregister-EWSGuiTenantData
```