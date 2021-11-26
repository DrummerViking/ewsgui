Function Method11 {
    <#
    .SYNOPSIS
    Method to get user's OOF Settings.
    
    .DESCRIPTION
    Method to get user's OOF Settings.
    
    .EXAMPLE
    PS C:\> Method11
    Method to get user's OOF Settings.

    #>
    [CmdletBinding()]
    param(
        # Parameters
    )
    $statusBarLabel.text = "Running..."
    $array = New-Object System.Collections.ArrayList
    $output = $service.GetUserOofSettings($email) | Select-Object `
        State, `
        ExternalAudience, `
        @{ Name = "StartTime" ; Expression = { $service.GetUserOofSettings($email).Duration.StartTime.ToString() } }, `
        @{ Name = "EndTime" ; Expression = { $service.GetUserOofSettings($email).Duration.EndTime.ToString() } }, `
        @{ Name = "InternalReply" ; Expression = { $service.GetUserOofSettings($email).InternalReply.Message } }, `
        @{ Name = "ExternalReply" ; Expression = { $service.GetUserOofSettings($email).ExternalReply.Message } }, `
        AllowExternalOof
    $array.Add($output)

    $dgResults.datasource = $array
    $dgResults.AutoResizeColumns()
    $dgResults.Visible = $True
    $txtBoxResults.Visible = $False
    $PremiseForm.refresh()
    $statusBarLabel.text = "Ready..."
    Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 11"
}