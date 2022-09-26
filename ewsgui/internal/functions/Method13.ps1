Function Method13 {
    <#
    .SYNOPSIS
    Get user's Delegates information
    
    .DESCRIPTION
    Get user's Delegates information
    
    .PARAMETER ClientID
    String parameter with the ClientID (or AppId) of your AzureAD Registered App.

    .PARAMETER TenantID
    String parameter with the TenantID your AzureAD tenant.

    .PARAMETER ClientSecret
    String parameter with the Client Secret which is configured in the AzureAD App.
    
    .EXAMPLE
    PS C:\> Method13
    Get user's Delegates information

    #>
    [CmdletBinding()]
    param(
        [String] $ClientID,

        [String] $TenantID,

        [String] $ClientSecret
    )
    $statusBarLabel.Text = "Running..."

    Test-StopWatch -Service $service -ClientID $ClientID -TenantID $TenantID -ClientSecret $ClientSecret

    # Create a mailbox object that represents the user in case we are impersonating.
    $mailbox = New-Object Microsoft.Exchange.WebServices.Data.Mailbox($email);
    # Call the GetDelegates method to get the delegates of the mailbox object.
    $delegates = $service.GetDelegates($mailbox , $true)
    $Collection = @()
    foreach( $Delegate in $delegates.DelegateUserResponses )
    {
        $Obj = "" | Select-Object EmailAddress,Inbox,Calendar,Contacts,Tasks,Notes,Journal,MeetingMessages,ViewPrivateItems
        $Obj.EmailAddress = $Delegate.DelegateUser.UserId.PrimarySmtpAddress
        $Obj.Inbox = $Delegate.DelegateUser.Permissions.InboxFolderPermissionLevel
        $Obj.Calendar = $Delegate.DelegateUser.Permissions.CalendarFolderPermissionLevel
        $Obj.Contacts = $Delegate.DelegateUser.Permissions.ContactsFolderPermissionLevel
        $Obj.Tasks = $Delegate.DelegateUser.Permissions.TasksFolderPermissionLevel
        $Obj.Notes = $Delegate.DelegateUser.Permissions.NotesFolderPermissionLevel
        $Obj.Journal = $Delegate.DelegateUser.Permissions.JournalFolderPermissionLevel
        $Obj.ViewPrivateItems = $Delegate.DelegateUser.ViewPrivateItems
        $Obj.MeetingMessages = $Delegate.DelegateUser.ReceiveCopiesOfMeetingMessages
        $Collection += $Obj
    }
    $array = New-Object System.Collections.ArrayList
    [int]$i = 0
    foreach ( $Del in $Collection )
    {
        $i++
        $output = $Del | Select-Object EmailAddress, Inbox, Calendar, Tasks, Notes, Journal, ViewPrivateItems, MeetingMessages
        $array.Add($output)
    }
    $dgResults.datasource = $array
    $dgResults.AutoResizeColumns()
    $dgResults.Visible = $True
    $txtBoxResults.Visible = $False
    $PremiseForm.refresh()
    $statusBarLabel.text = "Ready. Amount of Delegates: $i"
    Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 13" -Target $email
}