Function Method9 {
    <#
    .SYNOPSIS
    Method to get user's Inbox Rules.
    
    .DESCRIPTION
    Method to get user's Inbox Rules.
    
    .PARAMETER ClientID
    String parameter with the ClientID (or AppId) of your AzureAD Registered App.

    .PARAMETER TenantID
    String parameter with the TenantID your AzureAD tenant.

    .PARAMETER ClientSecret
    String parameter with the Client Secret which is configured in the AzureAD App.
    
    .EXAMPLE
    PS C:\> Method9
    Method to get user's Inbox Rules.

    #>
    [CmdletBinding()]
    param(
        [String] $ClientID,

        [String] $TenantID,

        [String] $ClientSecret
    )
    $statusBarLabel.Text = "Running..."

    Test-StopWatch -Service $service -ClientID $ClientID -TenantID $TenantID -ClientSecret $ClientSecret

    $txtBoxResults.Text = "This method is still under construction."
    $dgResults.Visible = $False
    $txtBoxResults.Visible = $True
    $PremiseForm.refresh()
    $statusBarLabel.text = "Ready..."
    
    <#
    $rules = $service.GetInboxRules()
    $array = New-Object System.Collections.ArrayList
    foreach ( $rule in $rules )
    {
        $output = $rule | select DisplayName, Conditions, Actions, Exceptions
        $array.Add($output)
    }
    $dgResults.datasource = $array
    $dgResults.AutoResizeColumns()
    $dgResults.Visible = $True
    $txtBoxResults.Visible = $False
    $PremiseForm.refresh()
    $statusBarLabel.text = "Ready..."
    Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 9" -Target $email
    #>
}