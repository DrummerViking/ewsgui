Function Method10 {
    <#
    .SYNOPSIS
    Method to get user's OOF Settings.
    
    .DESCRIPTION
    Method to get user's OOF Settings.
    
    .PARAMETER ClientID
    String parameter with the ClientID (or AppId) of your AzureAD Registered App.

    .PARAMETER TenantID
    String parameter with the TenantID your AzureAD tenant.

    .PARAMETER ClientSecret
    String parameter with the Client Secret which is configured in the AzureAD App.

    .EXAMPLE
    PS C:\> Method10
    Method to get user's OOF Settings.

    #>
    [CmdletBinding()]
    param(
        [String] $ClientID,

        [String] $TenantID,

        [String] $ClientSecret
    )
    $statusBarLabel.Text = "Running..."

    Test-StopWatch -Service $service -ClientID $ClientID -TenantID $TenantID -ClientSecret $ClientSecret

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
    Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 10" -Target $email
}