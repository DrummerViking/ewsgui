Function Method9 {
    <#
    .SYNOPSIS
    Method to delete a specific folder in the user mailbox.
    
    .DESCRIPTION
    Method to delete a specific folder in the user mailbox with 3 different deletion methods.
    
    .EXAMPLE
    PS C:\> Method9
    Method to delete a specific folder in the user mailbox.

    #>
    [CmdletBinding()]
    param(
        # Parameters
    )
    $statusBar.Text = "Running..."
        if ( $txtBoxFolderID.Text -ne "" )
        {    
            $sourceFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId($txtBoxFolderID.Text)
            $SourceFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$sourceFolderId)
            $sourceFolder.Delete($ComboOption)

            Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 9"
            $statusBar.Text = "Ready..."
            $PremiseForm.Refresh()
        }
        else
        {
            [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            $statusBar.Text = "Process finished with warnings/errors"
        }
}