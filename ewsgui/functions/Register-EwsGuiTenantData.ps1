function Register-EWsGuiTenantData {
    [CmdletBinding()]
    param (
        [String] $ClientID,

        [String] $TenantID,

        [String] $ClientSecret
    )
    
    begin {
        if ( $ClientID -eq '' -or $TenantID -eq '' -or $CertificateThumbprint -eq '' ) { 
            throw "Either ClientID, TenantID or ClientSecret are null or empty."
        }
    }
    
    process {
        Write-PSFMessage -Level Important -Message "Registering ClientID string to EWSGui Module."
        Set-PSFConfig -Module EwsGui -Name "ClientID" -Value $ClientID -Description "AppID of your Azure Registered App" -AllowDelete -PassThru | Register-PSFConfig
        
        Write-PSFMessage -Level Important -Message "Registering TenantID string to EWSGui Module."
        Set-PSFConfig -Module EwsGui -Name "TenantID" -Value $TenantID -Description "TenantID where your Azure App is registered." -AllowDelete -PassThru | Register-PSFConfig
        
        Write-PSFMessage -Level Important -Message "Registering ClientSecret string to EWSGui Module."
        Set-PSFConfig -Module EwsGui -Name "ClientSecret" -Value $clientSecret -Description "ClientSecret passcode for your Azure App" -AllowDelete -PassThru | Register-PSFConfig
    }
    
    end {
        
    }
}