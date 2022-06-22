function Unregister-EWsGuiTenantData {
    [CmdletBinding()]
    param (
        # Parameters
    )
    
    begin {

    }
    
    process {
        Write-PSFMessage -Level Important -Message "Unregistering ClientID, TenantID and ClientSecret strings from EWSGui Module."
        Unregister-PSFConfig -Module EWSGui
        remove-PSFConfig -Module ewsgui -Name clientID -Confirm:$false
        remove-PSFConfig -Module ewsgui -Name tenantID -Confirm:$false
        remove-PSFConfig -Module ewsgui -Name ClientSecret -Confirm:$false
    }
    
    end {
        
    }
}