Function Method1to5 {
    <#
    .SYNOPSIS
    Method to list folders in the user mailbox.
    
    .DESCRIPTION
    Method to list folders in the user mailbox, showing Folder name, FolderId, Number of items, and number of subfolders.
    
    .PARAMETER ClientID
    String parameter with the ClientID (or AppId) of your AzureAD Registered App.

    .PARAMETER TenantID
    String parameter with the TenantID your AzureAD tenant.

    .PARAMETER ClientSecret
    String parameter with the Client Secret which is configured in the AzureAD App.

    .EXAMPLE
    PS C:\> Method1to5
    lists folders in the user mailbox.

    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "")]
    [CmdletBinding()]
    param(
        [String] $ClientID,

        [String] $TenantID,

        [String] $ClientSecret
    )
    $statusBarLabel.Text = "Running..."

    Test-StopWatch -Service $service -ClientID $ClientID -TenantID $TenantID -ClientSecret $ClientSecret

    Function Find-Subfolders {
        Param (
            $array,

            $ParentFolderId,

            $ParentDisplayname
        )
        $sourceFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId($ParentFolderId)
        $rootfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$sourceFolderId)

        $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)
        #$FolderView.Traversal = "Deep"
        
        $rootfolder.load()
        foreach ($folder in $rootfolder.FindFolders($FolderView) ) {
            $i++
            $DisplayName = "$ParentDisplayname\$($Folder.Displayname)"
            $output = $folder | Select-Object @{N = "Displayname" ; E = {$DisplayName}}, @{N = "TotalItemsCount"; E = { $_.TotalCount } }, @{N = "# of Subfolders"; E = { $_.ChildFolderCount } }, Id
            $array.Add($output)
            if ($folder.ChildFolderCount -gt 0) {
                #write-host "looking for subfolders under $($folder.displayname)" -ForegroundColor Green
                Find-Subfolders -ParentFolderId $folder.id -ParentDisplayname $Displayname -Array $array
            }
        }
    }

    if ($radiobutton1.Checked) { $Wellknownfolder = "MsgFolderRoot" }
    elseif ($radiobutton2.Checked) { $Wellknownfolder = "ArchiveMsgFolderRoot" }
    elseif ($radiobutton3.Checked) { $Wellknownfolder = "PublicFoldersRoot" }
    elseif ($radiobutton4.Checked) { $Wellknownfolder = "RecoverableItemsRoot" }
    elseif ($radiobutton5.Checked) { $Wellknownfolder = "ArchiveRecoverableItemsRoot" }

    #listing all available folders in the mailbox
    $rootfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$Wellknownfolder)
    $array = New-Object System.Collections.ArrayList
    Find-Subfolders -ParentFolderId $rootfolder.id -Array $array -ParentDisplayname ""

    $dgResults.datasource = $array
    $dgResults.AutoResizeColumns()
    $dgResults.Visible = $True
    $txtBoxResults.Visible = $False
    $PremiseForm.refresh()
    $statusBarLabel.Text = "Ready. Folders found: $($array.Count)"
    Write-PSFMessage -Level Output -Message "Task finished succesfully" -FunctionName "Method 1-5" -Target $email
}