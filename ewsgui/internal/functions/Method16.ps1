Function Method16 {
    <#
    .SYNOPSIS
    Method to switch to another mailbox.
    
    .DESCRIPTION
    Method to switch to another mailbox.
    
    .PARAMETER ClientID
    String parameter with the ClientID (or AppId) of your AzureAD Registered App.

    .PARAMETER TenantID
    String parameter with the TenantID your AzureAD tenant.

    .PARAMETER ClientSecret
    String parameter with the Client Secret which is configured in the AzureAD App.
    
    .EXAMPLE
    PS C:\> Method16
    Method to switch to another mailbox.

    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidGlobalVars", "")]
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
        $TargetSmtpAddress = $txtBoxFolderID.Text
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetSmtpAddress)
        $service.HttpHeaders.Clear()
        $service.HttpHeaders.Add("X-AnchorMailbox", $TargetSmtpAddress)
        $Global:email = $TargetSmtpAddress

        $labImpersonation.Location = New-Object System.Drawing.Point(595,200)
        $labImpersonation.Size = New-Object System.Drawing.Size(300,20)
        $labImpersonation.Name = "labImpersonation"
        $labImpersonation.ForeColor = "Blue"
        $PremiseForm.Controls.Add($labImpersonation)
        $labImpersonation.Text = $Global:email
        $PremiseForm.Text = "Managing user: " + $Global:email + ". Choose your Option"

        Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 16" -Target $email
        $statusBarLabel.text = "Ready..."
        $PremiseForm.Refresh()
    }
    else
    {
        [Microsoft.VisualBasic.Interaction]::MsgBox("Email Address textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        $statusBarLabel.text = "Process finished with warnings/errors"
    }
}