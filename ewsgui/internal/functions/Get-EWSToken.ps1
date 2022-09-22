function Get-EWSToken {
    <#
    .SYNOPSIS
    Function to fetch an authentication AccessToken, or a RefreshToken.
    
    .DESCRIPTION
    Function to fetch an authentication AccessToken, or a RefreshToken.
    
    .PARAMETER Refresh
    Use this optional parameter to request a Refresh Token.

    .PARAMETER ClientID
    String parameter with the ClientID (or AppId) of your AzureAD Registered App.

    .PARAMETER TenantID
    String parameter with the TenantID your AzureAD tenant.

    .PARAMETER ClientSecret
    String parameter with the Client Secret which is configured in the AzureAD App.
    
    .EXAMPLE
    PS C:\> Get-EWSToken -ClientID "1234" -TenantID "abcd" -ClientSecret "a1b2c3d4"

    The function gets an authentication token based on the parameters passed.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidGlobalVars", "")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText", "")]
    [CmdletBinding()]
    param (
        [Switch] $Refresh = $false,

        [String] $ClientID,

        [String] $TenantID,

        [String] $ClientSecret
    )
    
    begin {
        
    }
    
    process {
        # creating time watcher or restarting if already passed 50 minutes since launch
        if ( $global:stopWatch.IsRunning -eq $true ) {
            $global:stopWatch.Restart()
        }
        else {
            $global:stopWatch = [System.Diagnostics.Stopwatch]::StartNew()
        }
        
        #Getting oauth credentials
        if ( $Refresh ) {
            Write-PSFMessage -Level Important -Message "Obtaining resfresh token to continue."
        }
        #region Connecting using Oauth with Application permissions with passed parameters
        if ( -not[String]::IsNullOrEmpty($ClientID) -or -not[String]::IsNullOrEmpty($TenantID) -or -not[String]::IsNullOrEmpty($ClientSecret) ) {
            Write-PSFMessage -Level Important -Message "Connecting using Oauth with Application permissions with passed parameters"
            $cid = $ClientID
            $tid = $TenantID
            $cs = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force

            $scopes = New-Object System.Collections.Generic.List[string]
            $scopes.Add("https://outlook.office365.com/.default")
            try {
                $token = Get-MsalToken -ClientId $cid -TenantId $tid -ClientSecret $cs -Scopes $scopes -ForceRefresh:$Refresh -ErrorAction Stop
            }
            catch {
                Write-PSFMessage -Level Error -Message "Something failed to get authentication token." -ErrorRecord $_
            }
            Write-PSFMessage -Level Important -Message "Connected using Application permissions with passed ClientID, TenantID and ClientSecret"
        }
        #endregion
        #region Connecting using Oauth with Application permissions with saved values in the module
        elseif (
            $null -ne (Get-PSFConfig -Module EwsGui -Name ClientID).value -and `
                $null -ne (Get-PSFConfig -Module EwsGui -Name TenantID).value -and `
                $null -ne (Get-PSFConfig -Module EwsGui -Name ClientSecret).value
        ) {
            Write-PSFMessage -Level Important -Message "Connecting using Oauth with Application permissions with saved values in the module"
            $cid = (Get-PSFConfig -Module EwsGui -Name ClientID).value
            $tid = (Get-PSFConfig -Module EwsGui -Name TenantID).value
            $cs = ConvertTo-SecureString -String (Get-PSFConfig -Module EwsGui -Name ClientSecret).value -AsPlainText -Force

            $scopes = New-Object System.Collections.Generic.List[string]
            $scopes.Add("https://outlook.office365.com/.default")
            try {
                $token = Get-MsalToken -ClientId $cid -TenantId $tid -ClientSecret $cs -Scopes $scopes -ForceRefresh:$Refresh -ErrorAction Stop
            }
            catch {
                Write-PSFMessage -Level Error -Message "Something failed to get authentication token." -ErrorRecord $_
            }
            Write-PSFMessage -Level Important -Message "Connected using Application permissions with registered ClientID, TenantID and ClientSecret embedded to the module."
        }
        #endregion
        #region Connecting using Oauth with delegated permissions
        else {
            Write-PSFMessage -Level Important -Message "Connecting using Oauth with delegated permissions"
            $scopes = New-Object System.Collections.Generic.List[string]
            $scopes.Add("https://outlook.office365.com/.default")
            try {
                $token = Get-MsalToken -ClientId "8799ab60-ace5-4bda-b31f-621c9f6668db" -RedirectUri "http://localhost/code" -Scopes $scopes -UseEmbeddedWebView -ForceRefresh:$Refresh -ErrorAction Stop
            }
            catch {
                if ( $_.Exception.toString().StartsWith("System.Threading.ThreadStateException: ActiveX control '8856f961-340a-11d0-a96b-00c04fd705a2'")) {
                    Write-PSFMessage -Level Error -Message "Known issue occurred. There is work in progress to fix authentication flow. More info at: https://github.com/agallego-css/ewsgui/issues/28"
                    Write-PSFMessage -Level Error -Message "Failed to obtain authentication token. Exiting script. Please rerun the script again and it should work."
                }
                else {
                    Write-PSFMessage -Level Error -Message "Something failed to get authentication token." -ErrorRecord $_
                }
                break
            }
            Write-PSFMessage -Level Important -Message "Connected using Delegated permissions with: $($token.Account.Username)"
        }
        #endregion
    }
    
    end {
        return $token
    }
}