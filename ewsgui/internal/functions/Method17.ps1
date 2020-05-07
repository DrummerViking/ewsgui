Function Method17 {
    <#
    .SYNOPSIS
    Method to switch to another mailbox.
    
    .DESCRIPTION
    Method to switch to another mailbox.
    
    .EXAMPLE
    PS C:\> Method17
    Method to switch to another mailbox.

    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidGlobalVars", "")]
    [CmdletBinding()]
    param(
        # Parameters
    )
    $statusBar.Text = "Running..."
    if ( $txtBoxFolderID.Text -ne "" )
    {
        $TargetSmtpAddress = $txtBoxFolderID.Text
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetSmtpAddress)
        $service.HttpHeaders.Clear()
        $service.HttpHeaders.Add("X-AnchorMailbox", $TargetSmtpAddress)
        $Global:email = $TargetSmtpAddress

        $labImpersonation.Location = New-Object System.Drawing.Point(575,231)
        $labImpersonation.Size = New-Object System.Drawing.Size(250,20)
        $labImpersonation.Name = "labImpersonation"
        $labImpersonation.ForeColor = "Blue"
        $PremiseForm.Controls.Add($labImpersonation)
        $labImpersonation.Text = $Global:email
        $PremiseForm.Text = "Managing user: " + $Global:email + ". Choose your Option"

        Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 17"
        $statusBar.Text = "Ready..."
        $PremiseForm.Refresh()
    }
    else
    {
        [Microsoft.VisualBasic.Interaction]::MsgBox("Email Address textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        $statusBar.Text = "Process finished with warnings/errors"
    }
}