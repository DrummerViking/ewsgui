Function Method8 {
    <#
    .SYNOPSIS
    Method to delete a specific folder in the user mailbox.
    
    .DESCRIPTION
    Method to delete a specific folder in the user mailbox with 3 different deletion methods.
    
    .PARAMETER ClientID
    String parameter with the ClientID (or AppId) of your AzureAD Registered App.

    .PARAMETER TenantID
    String parameter with the TenantID your AzureAD tenant.

    .PARAMETER ClientSecret
    String parameter with the Client Secret which is configured in the AzureAD App.

    .EXAMPLE
    PS C:\> Method8
    Method to delete a specific folder in the user mailbox.

    #>
    [CmdletBinding()]
    param(
        [String] $ClientID,

        [String] $TenantID,

        [String] $ClientSecret
    )
    $statusBarLabel.Text = "Running..."

    Test-StopWatch -Service $service -ClientID $ClientID -TenantID $TenantID -ClientSecret $ClientSecret

    if ( $txtBoxFolderID.Text -ne "" )
    {
        $sourceFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId($txtBoxFolderID.Text)
        $SourceFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$sourceFolderId)
        $sourceFolder.Delete($ComboOption)

        Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 8" -Target $email
        $statusBarLabel.text = "Ready..."
        $PremiseForm.Refresh()
    }
    else
    {
        [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        $statusBarLabel.text = "Process finished with warnings/errors"
    }
}