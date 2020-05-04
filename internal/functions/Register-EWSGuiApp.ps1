Function Register-EWSGuiApp {
    <#
    .SYNOPSIS
    Short description
    
    .DESCRIPTION
    Long description
    
    .EXAMPLE
    An example
    
    #>
    Invoke-PSFProtectedCommand -Action "Connecting to AzureAD" -Target "AzureAD" -ScriptBlock {
        Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to AzureAD"
        if ( !(Get-Module AzureAD -ListAvailable) -and !(Get-Module AzureAD) ) {
            Install-Module AzureAD -Force -ErrorAction Stop
        }
        try {
            Import-module AzureAD
            Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] We will connect to AzureAD to allow this app to connect to your tenant using OAUTH"
            $ConnStatus = Connect-AzureAD -ErrorAction Stop
        }
        catch {
            return $_
        }
    } -EnableException $true -PSCmdlet $PSCmdlet

    # register "PowerShellEWSScripts" as Enterprise App, by creating servicePrincipal (if not created)
    Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] Register `"PowerShellEWSScripts`" as Enterprise App, by creating servicePrincipal (if not created)"
    if ( $null -eq ( Get-AzureADServicePrincipal -All:$True | Where-object { $_.displayname -eq "PowerShellEWSScripts" } ) ) {
        $AzureADServicePrincipalParams = @{
            AccountEnabled            = $True
            AppId                     = "8799ab60-ace5-4bda-b31f-621c9f6668db"
            ServicePrincipalNames     = "8799ab60-ace5-4bda-b31f-621c9f6668db"
            AppRoleAssignmentRequired = $False
            DisplayName               = "PowerShellEWSScripts"
            PublisherName             = "Microsoft"
            ReplyUrls                 = "http://localhost/code"
            ServicePrincipalType      = "Application"
            Tags                      = "WindowsAzureActiveDirectoryIntegratedApp"
            ErrorAction               = "Stop"
        }
        $AppDetails = New-AzureADServicePrincipal @AzureADServicePrincipalParams
    }
    else {
        $AppDetails = Get-AzureADServicePrincipal -All:$True | Where-object { $_.displayname -eq "PowerShellEWSScripts" }
    }

    # register Service Principal Assignment between Global admin and the registered app:
    $AdminObjectId = (Get-AzureAdUser -Filter "userprincipalname eq '$($ConnStatus.Account.id)'").ObjectId
    if ( -not (Get-AzureADServiceAppRoleAssignment -ObjectId $AppDetails.ObjectId -All:$true | Where-Object { $_.PrincipalId -eq $AdminObjectId }) ) {
        $AppRoleAssignment = $appDetails | New-AzureADServiceAppRoleAssignment -PrincipalId $AdminObjectId -ResourceId $AppDetails.ObjectId -id "00000000-0000-0000-0000-000000000000" -ErrorAction Stop
    }

    # Grant consent to the App to access EXO and Windows Azure AD to sign in on behalf of the whole tenant (no user consent will be needed afterwards)
    Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] Grant consent to the App to access EXO and AzureAD to sign-in on behalf of the users"
    if ( -not (Get-AzureADServicePrincipalOAuth2PermissionGrant -ObjectId $AppDetails.ObjectId -All:$True | Where-Object { ($_.Scope -eq "User.Read" -and $_.ResourceID -eq "525065a3-df7d-474d-be08-90fa1d62d4bb") -or ($_.Scope -eq "EWS.AccessAsUser.All" -and $_.ResourceID -eq "8b951d63-7dd0-46e8-a326-15e3f6a26353") }) ) {
        $TenantId = (Get-AzureADTenantDetail).objectid
        $context = Get-AzContext
        $refreshToken = @($context.TokenCache.ReadItems() | Where-Object { $_.tenantId -eq $tenantId -and $_.ExpiresOn -gt (Get-Date) })[0].RefreshToken
        $body = "grant_type=refresh_token&refresh_token=$($refreshToken)&resource=74658136-14ec-4630-ad9b-26e160ff0fc6"
        $apiToken = Invoke-RestMethod "https://login.windows.net/$tenantId/oauth2/token" -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        $header = @{
            'Authorization'          = 'Bearer ' + $apiToken.access_token
            'X-Requested-With'       = 'XMLHttpRequest'
            'x-ms-client-request-id' = [guid]::NewGuid()
            'x-ms-correlation-id'    = [guid]::NewGuid()
        }
        $url = "https://main.iam.ad.ext.azure.com/api/RegisteredApplications/$($AppDetails.AppId)/Consent?onBehalfOfAll=true"
        Invoke-RestMethod –Uri $url –Headers $header –Method POST -ErrorAction Stop
    }
}