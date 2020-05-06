Function Method1to6 {
    <#
    .SYNOPSIS
    Method to list folders in the user mailbox.
    
    .DESCRIPTION
    Method to list folders in the user mailbox, showing Folder name, FolderId, Number of items, and number of subfolders.
    
    .EXAMPLE
    PS C:\> Method1to6
    lists folders in the user mailbox.

    #>
    [CmdletBinding()]
    param(
        # Parameters
    )
    $statusBar.Text = "Running..."
        if($radiobutton1.Checked){$Wellknownfolder = "MsgFolderRoot"}
        elseif($radiobutton2.Checked){$Wellknownfolder = "ArchiveMsgFolderRoot"}
        elseif($radiobutton3.Checked){$Wellknownfolder = "PublicFoldersRoot"}
        elseif($radiobutton5.Checked){$Wellknownfolder = "RecoverableItemsRoot"}
        elseif($radiobutton6.Checked){$Wellknownfolder = "ArchiveRecoverableItemsRoot"}
        elseif($radiobutton4.Checked){$Wellknownfolder = $txtBoxFolderID.Text}

        #listing all available folders in the mailbox
        $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100);
        if($radiobutton4.Checked){
            $sourceFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId($Wellknownfolder)
            $rootfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$sourceFolderId)
            }else{
            $rootfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$Wellknownfolder)
            }
        
        $rootfolder.load()
        $array = New-Object System.Collections.ArrayList
        foreach ($folder in $rootfolder.FindFolders($FolderView) )
        {
            $i++
            $output = $folder | Select-Object DisplayName, @{N="TotalItemsCount";E={$_.TotalCount}}, @{N="# of Subfolders";E={$_.ChildFolderCount}},Id
            $array.Add($output)
            }
        $dgResults.datasource = $array
        $dgResults.AutoResizeColumns()
        $dgResults.Visible = $True
        $txtBoxResults.Visible = $False
        $PremiseForm.refresh()
        $statusBar.Text = "Ready. Folders found: $i"
        Write-PSFMessage -Level Output -Message "Task finished succesfully" -FunctionName "Method 1-6"
}