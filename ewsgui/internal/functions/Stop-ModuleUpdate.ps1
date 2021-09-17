Function Stop-ModuleUpdate {
    <#
    .SYNOPSIS
    Function to stop checking for updates on this module and clear runspaces.
    
    .DESCRIPTION
    Function to stop checking for updates on this module and clear runspaces.
    
    .PARAMETER RunspaceData
    Runspace data retrieved from intial Start-ModuleUpdate function.

    .PARAMETER Confirm
    If this switch is enabled, you will be prompted for confirmation before executing any operations that change state.

    .PARAMETER WhatIf
    If this switch is enabled, no actions are performed but informational messages will be displayed that explain what would happen if the command were to run.
    
    .EXAMPLE
    PS C:\> Stop-ModuleUpdate -RunspaceData $data
    Runs the function to stop checking for update on this module and clear runspaces.
    #>
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'Low')]
    Param(
        $RunspaceData
    )
    # Receive Results and cleanup
	$null = $RunspaceData.Pipe.EndInvoke($RunspaceData.Status)
	$RunspaceData.Pipe.Dispose()

	# Cleanup Runspace Pool
	$RunspaceData.pool.Close()
	$RunspaceData.pool.Dispose()
}