Function Method12 {
    <#
    .SYNOPSIS
    Method to move items between folders.
    
    .DESCRIPTION
    Method to move items between folders by using FolderID values.
    
    .EXAMPLE
    PS C:\> Method12
    Method to move items between folders.

    #>
    [CmdletBinding()]
    param(
        # Parameters
    )
    $statusBar.Text = "Running..."
    if ( $txtBoxFolderID.Text -ne "")
    {
        # Creating Filter variables
        $FolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId($txtBoxFolderID.Text)
        $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$FolderID)
        $TargetFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId($txtBoxTargetFolderID.Text)
        $StartDate = $FromDatePicker.Value
        $EndDate = $ToDatePicker.Value
        $MsgSubject = $txtBoxSubject.text
        [int]$i = 0

        # Combining Filters into a single Collection
        $filters = @()
        if ( $MsgSubject -ne "" )
        {
            $Filter1 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject,$MsgSubject, [Microsoft.Exchange.WebServices.Data.ContainmentMode]::ExactPhrase, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)
            $filters += $Filter1
        }
        if ( $StartDate -ne "" )
        {
            $Filter2 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[DateTime]$StartDate)
            $filters += $Filter2
        }
        if ( $EndDate -ne "" )
        {
            $Filter3 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[DateTime]$EndDate)
            $filters += $Filter3
        }

        $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::AND,$filters)

        if ( $filters.Length -eq 0 )
        {
            $searchFilter = $Null
        }

        $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(250)
        $fiItems = $null
        $array = New-Object System.Collections.ArrayList
        do{
            $fiItems = $service.FindItems($Folder.Id, $searchFilter, $ivItemView)
            foreach ( $Item in $fiItems.Items )
            {
                $i++
                $output = $Item | Select-Object @{Name="Action";Expression={"Moving Item"}}, DateTimeReceived, Subject
                $array.Add($output)
                $tempItem = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service,$Item.Id)
                $tempItem.Move($TargetFolderId) | Out-Null
            }
            $ivItemView.Offset += $fiItems.Items.Count
            Start-Sleep -Milliseconds 500
        } while ( $fiItems.MoreAvailable -eq $true )
        $dgResults.datasource = $array
        $dgResults.AutoResizeColumns()
        $dgResults.Visible = $True
        $txtBoxResults.Visible = $False
        $PremiseForm.refresh()
        $statusBar.Text = "Ready. Moved Items: $i"
        Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 12"
    }
    else
    {
        [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        $statusBar.Text = "Process finished with warnings/errors"
    }
}