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
    $statusBarLabel.text = "Running..."
        if ( $txtBoxFolderID.Text -ne "" )
        {
            $sourceFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId($txtBoxFolderID.Text)
            $SourceFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$sourceFolderId)
            $sourceFolder.Delete($ComboOption)

            Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 9"
            $statusBarLabel.text = "Ready..."
            $PremiseForm.Refresh()
        }
        else
        {
            [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            $statusBarLabel.text = "Process finished with warnings/errors"
        }
}