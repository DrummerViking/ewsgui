Function Register-EWSGuiApp {
    <#
    .SYNOPSIS
    Script to create the Azure App Registration for EWSGui.

    .DESCRIPTION
    Script to create the Azure App Registration for EWSGui.
    It will require an additional PS module "Microsoft.Graph.Applications", if not already installed it will download it.
    You have to pass the list of app permissions you want to grant.
    You can use the "UseClientSecret" switch parameter to configure a new ClientSecret for the app. If this parameter is ommitted, we will use a Certificate.
    You can pass a certificate path if you have an existing certificate, or leave the parameter blank and a new self-signed certificate will be created.

    .PARAMETER AppName
    The friendly name of the app registration. By default will be "EWSGui Registered App".

    .PARAMETER TenantId
    Optional parameter to set the TenantID GUID.

    .PARAMETER StayConnected
    Use this optional parameter to not disconnect from Graph after the script execution.

    .PARAMETER ImportAppDataToModule
    Use this optional parameter to import your app's ClientId, TenantId and ClientSecret into the EWSGui module. In this way, the next time you run the app it will use the Application flow to authenticate with these values.

    .EXAMPLE
    PS C:\> Register-AzureADApp.ps1 -AppName "Graph DemoApp" -StayConnected

    The script will create a new AzureAD App Registration.
    The name of the app will be "Graph DemoApp".
    It will add the following API Permissions: "full_access_as_app".
    it will use a ClientSecret (later will be exposed).

    Once the app is created, the script will expose the link to grant "Admin consent" for the permissions requested.
    
    .NOTES
    General notes
#>
    [Cmdletbinding()]
    param(
        [Parameter(Mandatory = $false)]
        [String]
        $AppName = "EWSGui Registered App",

        [Parameter(Mandatory = $false)]
        [String]
        $TenantId,

        [Parameter(Mandatory = $false)]
        [Switch]
        $StayConnected,
    
        [Parameter(Mandatory = $false)]
        [Switch]
        $ImportAppDataToModule
    )
    # Required modules
    if ( -not(Get-module "Microsoft.Graph.Applications" -ListAvailable) ) {
        Install-Module "Microsoft.Graph.Applications" -Scope CurrentUser -Force
    }
    Import-Module "Microsoft.Graph.Applications"

    # Graph permissions variables
    $graphResourceId = "00000002-0000-0ff1-ce00-000000000000"
    $EwsApiPermission = @{
        Id   = "dc890d15-9560-4a4c-9b7f-a736ec74ec40" # "full_access_as_app"
        Type = "Role"
    }

    # Requires an admin
    if ($TenantId) {
        Connect-MgGraph -Scopes "Application.ReadWrite.All User.Read" -TenantId $TenantId
    }
    else {
        Connect-MgGraph -Scopes "Application.ReadWrite.All User.Read"
    }

    # Get context for access to tenant ID
    $context = Get-MgContext

    # Create app registration
    $appRegistration = New-MgApplication -DisplayName $AppName -SignInAudience "AzureADMyOrg" `
        -Web @{ RedirectUris = "http://localhost"; } `
        -RequiredResourceAccess @{ ResourceAppId = $graphResourceId; ResourceAccess = @($EwsApiPermission) } `
        -AdditionalProperties @{}

    $appObjId = Get-MgApplication -Filter "AppId eq '$($appRegistration.Appid)'"
    $passwordCred = @{
        displayName = 'Secret created in PowerShell'
        endDateTime = (Get-Date).Addyears(1)
    }
    $secret = Add-MgApplicationPassword -applicationId $appObjId.Id -PasswordCredential $passwordCred
    Write-Host "App registration created with app ID $($appRegistration.AppId) and clientSecret: $($secret.SecretText)" -ForegroundColor Cyan
    Write-Host -ForegroundColor Cyan "Please take note of your client secret as it will not be shown anymore"
    # Create corresponding service principal
    New-MgServicePrincipal -AppId $appRegistration.AppId -AdditionalProperties @{} | Out-Null
    Write-Host -ForegroundColor Cyan "Service principal created"
    Write-Host
    Write-Host -ForegroundColor Green "Success"
    Write-Host

    # Generate admin consent URL
    $adminConsentUrl = "https://login.microsoftonline.com/" + $context.TenantId + "/adminconsent?client_id=" + $appRegistration.AppId
    Write-Host -ForeGroundColor Yellow "Please go to the following URL in your browser to provide admin consent"
    Write-Host $adminConsentUrl
    Write-Host

    if ( $ImportAppDataToModule ) {
        Import-EWsGuiAADAppData -ClientID $appRegistration.AppId -TenantID $context.TenantId -ClientSecret $secret.SecretText
    }

    if ($StayConnected -eq $false) {
        Write-Host
        $null = Disconnect-MgGraph
        Write-Host "Disconnected from Microsoft Graph"
    }
    else {
        Write-Host
        Write-Host -ForegroundColor Yellow "The connection to Microsoft Graph is still active. To disconnect, use Disconnect-MgGraph"
    }
}