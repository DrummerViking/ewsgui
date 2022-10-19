function Remove-EWsGuiAADAppData {
    <#
    .SYNOPSIS
    Function to remove ClientID, TenantID and ClientSecret to the EWSGui powershell module.
    
    .DESCRIPTION
    Function to remove ClientID, TenantID and ClientSecret to the EWSGui powershell module.
    
    .PARAMETER Confirm
    If this switch is enabled, you will be prompted for confirmation before executing any operations that change state.

    .PARAMETER WhatIf
    If this switch is enabled, no actions are performed but informational messages will be displayed that explain what would happen if the command were to run.

    .EXAMPLE
    PS C:\> Remove-EWsGuiAADAppData

    The script will Remove these values in the EWSGui module to be used automatically.
    #>
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'Low')]
    param (
        # Parameters
    )
    
    begin {

    }
    
    process {
        Write-PSFMessage -Level Important -Message "Removing ClientID, TenantID and ClientSecret strings from EWSGui Module."
        Unregister-PSFConfig -Module EWSGui
        remove-PSFConfig -Module ewsgui -Name clientID -Confirm:$false
        remove-PSFConfig -Module ewsgui -Name tenantID -Confirm:$false
        remove-PSFConfig -Module ewsgui -Name ClientSecret -Confirm:$false
    }
    
    end {
        
    }
}