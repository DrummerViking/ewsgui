Function Method7 {
    <#
    .SYNOPSIS
    Method to create a custom folder in mailbox's Root.
    
    .DESCRIPTION
    Method to create a custom folder in mailbox's Root.
    
    .PARAMETER ClientID
    String parameter with the ClientID (or AppId) of your AzureAD Registered App.

    .PARAMETER TenantID
    String parameter with the TenantID your AzureAD tenant.

    .PARAMETER ClientSecret
    String parameter with the Client Secret which is configured in the AzureAD App.
    
    .EXAMPLE
    PS C:\> Method7
    Method to create a custom folder in mailbox's Root.

    #>
    [CmdletBinding()]
    param(
        [String] $ClientID,

        [String] $TenantID,

        [String] $ClientSecret
    )
    
    if ( $txtBoxFolderID.Text -ne "" )
    {
        $statusBarLabel.text = "Running..."
        Test-StopWatch -Service $service -ClientID $ClientID -TenantID $TenantID -ClientSecret $ClientSecret

        $folder = new-object Microsoft.Exchange.WebServices.Data.Folder($service)
        $folder.DisplayName = $txtBoxFolderID.Text
        $folder.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)

        Write-PSFMessage -Level Host -Message "Task finished succesfully. Folder Created: $($txtBoxFolderID.Text)" -FunctionName "Method 7" -Target $email
        $statusBarLabel.text = "Ready..."
        $PremiseForm.Refresh()
    }
    else
    {
        [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        $statusBarLabel.text = "Method 8 finished with warnings/errors"
    }
}