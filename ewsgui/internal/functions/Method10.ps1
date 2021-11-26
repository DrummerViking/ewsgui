Function Method10 {
    <#
    .SYNOPSIS
    Method to get user's Inbox Rules.
    
    .DESCRIPTION
    Method to get user's Inbox Rules.
    
    .EXAMPLE
    PS C:\> Method10
    Method to get user's Inbox Rules.

    #>
    [CmdletBinding()]
    param(
        # Parameters
    )
    $statusBarLabel.text = "Running..."
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
    Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 10"
    #>
}