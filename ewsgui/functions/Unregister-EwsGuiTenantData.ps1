function Unregister-EWsGuiAADAppData {
    <#
    .SYNOPSIS
    Function to unregister ClientID, TenantID and ClientSecret to the EWSGui powershell module.
    
    .DESCRIPTION
    Function to unregister ClientID, TenantID and ClientSecret to the EWSGui powershell module.
    
    .EXAMPLE
    PS C:\> Unregister-EWsGuiAADAppData

    The script will unregister these values in the EWSGui module to be used automatically.
    #>
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