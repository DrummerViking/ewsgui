function Test-StopWatch {
    <#
    .SYNOPSIS
    Function to check time elapsed since the last token was retrieved.
    If the time elapsed is greater than 50 minutes, we will go and fetch a new AccessToken based on the refresh token.
    
    .DESCRIPTION
    Function to check time elapsed since the last token was retrieved.
    If the time elapsed is greater than 50 minutes, we will go and fetch a new AccessToken based on the refresh token.

    .PARAMETER service
    EWS Service object.

    .PARAMETER ClientID
    String parameter with the ClientID (or AppId) of your AzureAD Registered App.

    .PARAMETER TenantID
    String parameter with the TenantID your AzureAD tenant.

    .PARAMETER ClientSecret
    String parameter with the Client Secret which is configured in the AzureAD App.
    
    .EXAMPLE
    PS C:\> Test-StopWatch -Service $svcObj -ClientID "1234" -TenantID "abcd" -ClientSecret "a1b2c3d4"

    The function will check the time elapsed in the stop watcher, and fetch a new AccessToken if needed.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidGlobalVars", "")]
    [CmdletBinding()]
    param (
        $service,

        [String] $ClientID,

        [String] $TenantID,

        [String] $ClientSecret
    )
    
    begin {
        
    }
    
    process {
        if ( $global:stopWatch.Elapsed.Minutes -gt 50) {
            $token = Get-EWSToken -ClientID $ClientID -TenantID $TenantID -ClientSecret $ClientSecret -Refresh
            $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($Token.AccessToken)
            $Service.Credentials = $exchangeCredentials
        }
    }
    
    end {
        
    }
}