Function Method8 {
    <#
    .SYNOPSIS
    Method to create a custom folder in mailbox's Root.
    
    .DESCRIPTION
    Method to create a custom folder in mailbox's Root.
    
    .EXAMPLE
    PS C:\> Method8
    Method to create a custom folder in mailbox's Root.

    #>
    [CmdletBinding()]
    param(
        # Parameters
    )
    if ( $txtBoxFolderID.Text -ne "" )
    {
        $statusBarLabel.text = "Running..."
        $folder = new-object Microsoft.Exchange.WebServices.Data.Folder($service)
        $folder.DisplayName = $txtBoxFolderID.Text
        $folder.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)

        Write-PSFMessage -Level Host -Message "Task finished succesfully. Folder Created: $($txtBoxFolderID.Text)" -FunctionName "Method 8"
        $statusBarLabel.text = "Ready..."
        $PremiseForm.Refresh()
    }
    else
    {
        [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        $statusBarLabel.text = "Method 8 finished with warnings/errors"
    }
}