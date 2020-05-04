﻿Function Start-EWSGui {
    <#
    .SYNOPSIS
        Allows to perform 16 different operations using EWS API.

    .DESCRIPTION
        Allows to perform 16 different operations using EWS API:
        1) List Folders in Root
        2) List Folders in Archive Root
        3) List Folders in Public Folder Root
        4) List subFolders from a desired Parent Folder
        5) List folders in Recoverable Items Root folder
        6) List folders in Recoverable Items folder in Archive
        7) List Items in a desired Folder
        8) Create a custom Folder in Root
        9) Delete a Folder
        10) Get user's Inbox Rules
        11) Get user's OOF Settings
        12) Move items between folders
        13) Delete a subset of items in a folder
        14) Get user's Delegate information
        15) Change sensitivity to items in a folder
        16) Remove OWA configurations
        17) Switch to another Mailbox

    .EXAMPLE
    PS C:\ Start-Ewsgui.ps1
    Runs the GUI tool to use EWS with Exchange server and Online.

    .NOTES
        Version History:
        1.00 - 01/30/2018 - Project start
        1.00 - 02/19/2018 - First Release
        1.70 - 03/13/2018 - Added new method "remove OWA Configurations"
                            Added $service.HttpHeaders.Add("X-AnchorMailbox", $TargetSmtpAddress) , for better Impersonation performance
                            Removed AutodiscoverSCPLookup, as most of the uses of this app is with Exchange Online
        1.80 - 04/16/2018 - Optimized logon options. If we choose 'Office 365', we will not use SCP and we hard-code EXO endpoint.
        1.82 - 02/15/2019 - Added 2 columns to folder lists methods 1-6 : TotalItemsCount, # of Subfolders
        2.00 - 04/29/2020 - Moving tool to Module in GitHub
    #>
    $script:nl = "`r`n"
    $ProgressPreference = "SilentlyContinue"

    function GenerateForm {
         
    #region Import the Assemblies
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName Microsoft.VisualBasic
    [System.Windows.Forms.Application]::EnableVisualStyles() 
    #endregion
     
    #region Generated Form Objects
    $Global:PremiseForm = New-Object System.Windows.Forms.Form
    $Global:radiobutton1 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton2 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton3 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton4 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton5 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton6 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton7 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton8 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton9 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton10 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton11 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton12 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton13 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton14 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton15 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton16 = New-Object System.Windows.Forms.RadioButton
    $Global:radiobutton17 = New-Object System.Windows.Forms.RadioButton
    $Global:labImpersonation = New-Object System.Windows.Forms.Label
    $labImpersonationHelp = New-Object System.Windows.Forms.Label
    $Global:buttonGo = New-Object System.Windows.Forms.Button
    $Global:buttonExit = New-Object System.Windows.Forms.Button

    $Global:dgResults = New-Object System.Windows.Forms.DataGridView 
    $Global:txtBoxResults = New-Object System.Windows.Forms.Label
    $Global:InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
    [String]$Global:email = $null
    #endregion Generated Form Objects
     
    #region Processes

    #region Process1to6
    $Process1to6 = {
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
        write-host "process1-6 finished" -ForegroundColor Yellow
        
    }
    #endregion Process1to6

    #region Process7
    $Process7 = {
        $statusBar.Text = "Running..."
        if($txtBoxFolderID.Text -ne ""){
            # Creating Filter variables
            
            $FolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId($txtBoxFolderID.Text)
            $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$FolderID)
            $StartDate = $FromDatePicker.Value
            $EndDate = $ToDatePicker.Value
            $MsgSubject = $txtBoxSubject.text
            
            # Combining Filters into a single Collection
            $filters = @()
            if($MsgSubject -ne ""){
                $Filter1 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject,$MsgSubject, [Microsoft.Exchange.WebServices.Data.ContainmentMode]::ExactPhrase, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)
                $filters += $Filter1
                }
            if($StartDate -ne ""){
                $Filter2 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[DateTime]$StartDate)
                $filters += $Filter2
                }
            if($EndDate -ne ""){
                $Filter3 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[DateTime]$EndDate)
                $filters += $Filter3
                }
                
            $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::AND,$filters)
            
            if($filters.Length -eq 0){
                $searchFilter = $Null
                }
     
            $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(250)  
         
            $fiItems = $null
            $array = New-Object System.Collections.ArrayList  
            do{  
                $fiItems = $service.FindItems($Folder.Id, $searchFilter, $ivItemView)  
                foreach($Item in $fiItems.Items){  
                    $i++
                    $output = $Item | Select-Object Subject, Sender, DateTimeReceived, Size #, Id
                    $array.Add($output)
                                
                }
                $ivItemView.Offset += $fiItems.Items.Count  
            }while($fiItems.MoreAvailable -eq $true)  
            $dgResults.datasource = $array
            $dgResults.AutoResizeColumns()
            $dgResults.Visible = $True
            $txtBoxResults.Visible = $False
            $PremiseForm.refresh()
            $statusBar.Text = "Ready. Items found: $i"
            write-host "Process7 finished" -ForegroundColor Yellow
        }else{
            [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            $statusBar.Text = "Process finished with warnings/errors"
        }
        
    }
    #endregion Process7

    #region Process8
    $Process8 = {
    if($txtBoxFolderID.Text -ne ""){
        $statusBar.Text = "Running..."
        $folder = new-object Microsoft.Exchange.WebServices.Data.Folder($service)
        $folder.DisplayName = $txtBoxFolderID.Text
        $folder.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)

        Write-Host "process8 finished. Folder Created:" $txtBoxFolderID.Text -ForegroundColor Yellow
        $statusBar.Text = "Ready..."
        $PremiseForm.Refresh()
        }else{
            [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            $statusBar.Text = "Process finished with warnings/errors"
        }
    }
    #endregion Process8

    #region Process9
    $Process9 = {
        $statusBar.Text = "Running..."
        if($txtBoxFolderID.Text -ne ""){
            
            $sourceFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId($txtBoxFolderID.Text)
            $SourceFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$sourceFolderId)
            $sourceFolder.Delete($ComboOption)

            Write-Host "process9 finished." -ForegroundColor Yellow
            $statusBar.Text = "Ready..."
            $PremiseForm.Refresh()
            }else{
            [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            $statusBar.Text = "Process finished with warnings/errors"
        }
        
    }
    #endregion Process9

    #region Process10
    $Process10 = {
        $statusBar.Text = "Running..."
        $txtBoxResults.Text = "This method is still under construction."
        $dgResults.Visible = $False
        $txtBoxResults.Visible = $True
        $PremiseForm.refresh()
        $statusBar.Text = "Ready..."
        
        <#
        $rules = $service.GetInboxRules()
        $array = New-Object System.Collections.ArrayList
        foreach ($rule in $rules)
        {
            $output = $rule | select DisplayName, Conditions, Actions, Exceptions
            $array.Add($output)
            }
        $dgResults.datasource = $array
        $dgResults.AutoResizeColumns()
        $dgResults.Visible = $True
        $txtBoxResults.Visible = $False
        $PremiseForm.refresh()
        $statusBar.Text = "Ready..."
        write-host "process10 finished" -ForegroundColor Yellow    
        #>
    }
    #endregion Process10

    #region Process11
    $Process11 = {
        $statusBar.Text = "Running..."
        $array = New-Object System.Collections.ArrayList
        $output = $service.GetUserOofSettings($email) | Select-Object State, ExternalAudience, @{Name="StartTime";Expression={$service.GetUserOofSettings($email).Duration.StartTime.ToString()}}, @{Name="EndTime";Expression={$service.GetUserOofSettings($email).Duration.EndTime.ToString()}}, @{Name="InternalReply";Expression={$service.GetUserOofSettings($email).InternalReply.Message}}, @{Name="ExternalReply";Expression={$service.GetUserOofSettings($email).ExternalReply.Message}}, AllowExternalOof
        $array.Add($output)
        
        $dgResults.datasource = $array
        $dgResults.AutoResizeColumns()
        $dgResults.Visible = $True
        $txtBoxResults.Visible = $False
        $PremiseForm.refresh()
        $statusBar.Text = "Ready..."
        write-host "process11 finished" -ForegroundColor Yellow        
    }
    #endregion Process11

    #region Process12
    $Process12 = {
        $statusBar.Text = "Running..."
        if($txtBoxFolderID.Text -ne ""){
            # Creating Filter variables
            $FolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId($txtBoxFolderID.Text)
            $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$FolderID)
            $TargetFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId($txtBoxTargetFolderID.Text)
            $StartDate = $FromDatePicker.Value
            $EndDate = $ToDatePicker.Value
            $MsgSubject = $txtBoxSubject.text
            
            # Combining Filters into a single Collection
            $filters = @()
            if($MsgSubject -ne ""){
                $Filter1 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject,$MsgSubject, [Microsoft.Exchange.WebServices.Data.ContainmentMode]::ExactPhrase, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)
                $filters += $Filter1
                }
            if($StartDate -ne ""){
                $Filter2 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[DateTime]$StartDate)
                $filters += $Filter2
                }
            if($EndDate -ne ""){
                $Filter3 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[DateTime]$EndDate)
                $filters += $Filter3
                }
                
            $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::AND,$filters)
            
            if($filters.Length -eq 0){
                $searchFilter = $Null
                }
     
            $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(250)  
            $fiItems = $null
            $array = New-Object System.Collections.ArrayList  
            do{  
                $fiItems = $service.FindItems($Folder.Id, $searchFilter, $ivItemView)  
                foreach($Item in $fiItems.Items){  
                    $i++
                    $output = $Item | Select-Object @{Name="Action";Expression={"Moving Item"}}, DateTimeReceived, Subject
                    $array.Add($output)
                    $tempItem = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service,$Item.Id)
                    $tempItem.Move($TargetFolderId) | Out-Null
                }
                $ivItemView.Offset += $fiItems.Items.Count
                Start-Sleep -Milliseconds 500  
            }while($fiItems.MoreAvailable -eq $true)  
            $dgResults.datasource = $array
            $dgResults.AutoResizeColumns()
            $dgResults.Visible = $True
            $txtBoxResults.Visible = $False
            $PremiseForm.refresh()
            $statusBar.Text = "Ready. Moved Items: $i"
            write-host "Process12 finished" -ForegroundColor Yellow
        }else{
            [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            $statusBar.Text = "Process finished with warnings/errors"
        }
        
    }
    #endregion Process12

    #region Process13
    $Process13 = {
        $statusBar.Text = "Running..."
        if($txtBoxFolderID.Text -ne ""){
            # Creating Filter variables
            $FolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId($txtBoxFolderID.Text)
            $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$FolderID)
            $StartDate = $FromDatePicker.Value
            $EndDate = $ToDatePicker.Value
            $MsgSubject = $txtBoxSubject.text
            
            # Combining Filters into a single Collection
            $filters = @()
            if($MsgSubject -ne ""){
                $Filter1 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject,$MsgSubject, [Microsoft.Exchange.WebServices.Data.ContainmentMode]::ExactPhrase, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)
                $filters += $Filter1
                }
            if($StartDate -ne ""){
                $Filter2 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[DateTime]$StartDate)
                $filters += $Filter2
                }
            if($EndDate -ne ""){
                $Filter3 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[DateTime]$EndDate)
                $filters += $Filter3
                }
                
            $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::AND,$filters)
            
            if($filters.Length -eq 0){
                $searchFilter = $Null
                }
     
            $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(250)  
         
            $fiItems = $null
            $array = New-Object System.Collections.ArrayList  
            do{  
                $fiItems = $service.FindItems($Folder.Id, $searchFilter, $ivItemView)  
                foreach($Item in $fiItems.Items){  
                    $i++
                    $output = $Item | Select-Object @{Name="Action";Expression={"Deleting Item"}}, DateTimeReceived, Subject
                    $array.Add($output)
                
                    $tempItem = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service,$Item.Id)
                    $tempItem.Delete($ComboOption, $True)
                }
                $ivItemView.Offset += $fiItems.Items.Count
                Start-Sleep -Milliseconds 500  
            }while($fiItems.MoreAvailable -eq $true)  
            $dgResults.datasource = $array
            $dgResults.AutoResizeColumns()
            $dgResults.Visible = $True
            $txtBoxResults.Visible = $False
            $PremiseForm.refresh()
            $statusBar.Text = "Ready. Deleted items: $i"
            write-host "Process13 finished" -ForegroundColor Yellow
        }else{
            [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            $statusBar.Text = "Process finished with warnings/errors"
        }
        
    }
    #endregion Process13

    #region Process14
    $Process14 = {
        $statusBar.Text = "Running..."
        #// Create a mailbox object that represents the user.
        $mailbox = New-Object Microsoft.Exchange.WebServices.Data.Mailbox($Global:email);
        #// Call the GetDelegates method to get the delegates of the primary user.
        $delegates = $service.GetDelegates($mailbox, $true);
        $Collection = @()
        foreach($Delegate in $delegates.DelegateUserResponses)
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
        foreach ($Del in $Collection)
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
        $statusBar.Text = "Ready. Amount of Delegates: $i"
        write-host "process14 finished" -ForegroundColor Yellow    
        
    }
    #endregion Process14

    #region Process15
    $Process15 = {
    $statusBar.Text = "Running..."
        if($txtBoxFolderID.Text -ne ""){
            # Creating Filter variables
            $FolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId($txtBoxFolderID.Text)
            $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$FolderID)
            $StartDate = $FromDatePicker.Value
            $EndDate = $ToDatePicker.Value
            $MsgSubject = $txtBoxSubject.text
            
            # Combining Filters into a single Collection
            $filters = @()
            if($MsgSubject -ne ""){
                $Filter1 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject,$MsgSubject, [Microsoft.Exchange.WebServices.Data.ContainmentMode]::ExactPhrase, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)
                $filters += $Filter1
                }
            if($StartDate -ne ""){
                $Filter2 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[DateTime]$StartDate)
                $filters += $Filter2
                }
            if($EndDate -ne ""){
                $Filter3 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[DateTime]$EndDate)
                $filters += $Filter3
                }
                
            $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::AND,$filters)
            
            if($filters.Length -eq 0){
                $searchFilter = $Null
                }
     
            $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(250)  
         
            $fiItems = $null
            $array = New-Object System.Collections.ArrayList  
            do{  
                $fiItems = $service.FindItems($Folder.Id, $searchFilter, $ivItemView)  
                foreach($Item in $fiItems.Items){  
                    $i++
                    $output = $Item | Select-Object @{Name="Action";Expression={"Applying Sensitivity" + $DeleteOpt}}, DateTimeReceived, Subject
                    $array.Add($output)

                    $tempItem = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service,$Item.Id)
                    $tempItem.Sensitivity = $DeleteOpt
                    $tempItem.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
                }
                $ivItemView.Offset += $fiItems.Items.Count
                Start-Sleep -Milliseconds 500  
            }while($fiItems.MoreAvailable -eq $true)  
            $dgResults.datasource = $array
            $dgResults.AutoResizeColumns()
            $dgResults.Visible = $True
            $txtBoxResults.Visible = $False
            $PremiseForm.refresh()
            $statusBar.Text = "Ready. Items changed: $i"
            write-host "Process15 finished" -ForegroundColor Yellow
        }else{
            [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            $statusBar.Text = "Process finished with warnings/errors"
        }
        
    }
    #endregion Process15

    #region Process16
    $Process16 = {
        $statusBar.Text = "Running..."
        $output = "Checking" + $ComboOption2
        $txtBoxResults.Text = $output
        $txtBoxResults.Visible = $True
        $PremiseForm.Refresh()

            $fid = $null
            if ($ComboOption1 -eq "Root")
            {
                $fid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root, $email) 
            }
            elseif ($ComboOption1 -eq "Calendar")
            {
            $fid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $email) 
            }
            elseif ($ComboOption1 -eq "Inbox")
            {
            $fid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $email) 
            }
        

            
            if($ComboOption2 -ne "CleanFinders"){
                try{
                    $Config = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($Service, $ComboOption2, $fid, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)
    
                    $Config.Delete();
                    $output = $output + $nl + "Deleted $ComboOption2"
                }
                catch{
                    $output = $output + $nl + "$ComboOption2 doesn't exist"
                }
                $statusBar.Text = "Ready..."
                Write-Host "Process16 finished." -ForegroundColor Yellow
                $txtBoxResults.Text = $output
                $txtBoxResults.Visible = $True
                $PremiseForm.Refresh()

            }else{
                # Creating folder object (SearchFolders also know as Finder)
                $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SearchFolders,$SmtpAddress)

                # Opening the bind to user Folder Finder
                $output = $output + $nl + "Opening Mailbox: $email"
                try{
                    $finderFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$folderid)
            
                    $output = $output + $nl + "Cleaning SearchFolder (same as Outlook /Cleanfinders)"

                    # If the bind was created clean the folder Finder
                    Try{
                        $finderFolder.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, $true)
                        $output = $output + $nl + "The Cleanup process for the Mailbox: $email Succeed!"
                    }
                    catch{
                        $output = $output + $nl + "Fail to clean Search folder Mailbox: $email"
                    }
                }
                catch{
                    $output = $output + $nl + "Fail to open Mailbox: $email"
                }
                $txtBoxResults.Text = $output
                $txtBoxResults.Visible = $True
                $statusBar.Text = "Ready..."
                Write-Host "Process16 finished." -ForegroundColor Yellow
                $PremiseForm.Refresh()
                
                #Cleaning Variables
                $SmtpAddress = $null
                $finderFolder = $null
                $folderid = $null
            }    
    }
    #endregion Process16

    #region Process17
    $Process17 = {
    $statusBar.Text = "Running..."
    if($txtBoxFolderID.Text -ne ""){
        $TargetSmtpAddress = $txtBoxFolderID.Text
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetSmtpAddress)
        $service.HttpHeaders.Add("X-AnchorMailbox", $TargetSmtpAddress)
        $Global:email = $TargetSmtpAddress

        $labImpersonation.Location = New-Object System.Drawing.Point(575,231)
        $labImpersonation.Size = New-Object System.Drawing.Size(250,20)
        $labImpersonation.Name = "labImpersonation"
        $labImpersonation.ForeColor = "Blue"
        $PremiseForm.Controls.Add($labImpersonation)
        $labImpersonation.Text = $Global:email
        $PremiseForm.Text = "Managing user: " + $Global:email + ". Choose your Option"
        
        Write-Host "Process17 finished." -ForegroundColor Yellow
        $statusBar.Text = "Ready..."
        $PremiseForm.Refresh()
        
    }else{
        [Microsoft.VisualBasic.Interaction]::MsgBox("Email Address textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        $statusBar.Text = "Process finished with warnings/errors"
        }
        
    }
    #endregion Process17

    #endregion Processes

    # Loading EWS API dll file
    Import-EWSDLL

    # registering EWS API as an Enterprise App in Azure AD
    # Register-EWSGuiApp
    
    # Connecting to EWS and creating service object
    $service = Connect-EWSService

    $ExpandFilters = {
    # Removing all controls, in order to reload the screen appropiately for each selection
    $PremiseForm.Controls.RemoveByKey("FromDate")
    $PremiseForm.Controls.RemoveByKey("FromDatePicker")
    $PremiseForm.Controls.RemoveByKey("ToDate")
    $PremiseForm.Controls.RemoveByKey("ToDatePicker") 
    $PremiseForm.Controls.RemoveByKey("labSubject")
    $PremiseForm.Controls.RemoveByKey("txtBoxSubject")
    $PremiseForm.Controls.RemoveByKey("labFolderID")
    $PremiseForm.Controls.RemoveByKey("txtBoxFolderID")
    $PremiseForm.Controls.RemoveByKey("txtBoxTargetFolderID")
    $PremiseForm.Controls.RemoveByKey("labTargetFolderID")
    $PremiseForm.Controls.RemoveByKey("labelCombobox")
    $PremiseForm.Controls.RemoveByKey("comboBoxMenu")
    $PremiseForm.Controls.RemoveByKey("labelComboboxFolder")
    $PremiseForm.Controls.RemoveByKey("comboBoxFolder")
    $PremiseForm.Controls.RemoveByKey("labelComboboxConfig")
    $PremiseForm.Controls.RemoveByKey("comboBoxConfig")

    $Global:labFromDate = New-Object System.Windows.Forms.Label
    $Global:FromDatePicker = New-Object System.Windows.Forms.DateTimePicker
    $Global:labToDate = New-Object System.Windows.Forms.Label
    $Global:ToDatePicker = New-Object System.Windows.Forms.DateTimePicker
    $Global:labSubject = New-Object System.Windows.Forms.Label
    $Global:txtBoxSubject = New-Object System.Windows.Forms.TextBox
    $Global:labFolderID = New-Object System.Windows.Forms.Label
    $Global:txtBoxFolderID = New-Object System.Windows.Forms.TextBox
    $Global:labTargetFolderID = New-Object System.Windows.Forms.Label
    $Global:txtBoxTargetFolderID = New-Object System.Windows.Forms.TextBox
    $Global:labelCombobox = New-Object System.Windows.Forms.Label
    $Global:comboBoxMenu = New-Object System.Windows.Forms.ComboBox
    $Global:labelComboboxFolder = New-Object System.Windows.Forms.Label
    $Global:comboBoxFolder = New-Object System.Windows.Forms.ComboBox
    $Global:labelComboboxConfig = New-Object System.Windows.Forms.Label
    $Global:comboBoxConfig = New-Object System.Windows.Forms.ComboBox

    #Label FromDate
    $labFromDate.Location = New-Object System.Drawing.Point(5,285)
    $labFromDate.Size = New-Object System.Drawing.Size(80,35)
    $labFromDate.Name = "FromDate"
    $labFromDate.Text = "From or greater than"
     
     
    # FromDate Date Picker
    $FromDatePicker.DataBindings.DefaultDataSourceUpdateMode = 0
    $FromDatePicker.Location = New-Object System.Drawing.Point(100,285)
    $FromDatePicker.Name = "FromDatePicker"
    $FromDatePicker.Text = ""


    #Label ToDate
    $labToDate.Location = New-Object System.Drawing.Point(5,330)
    $labToDate.Name = "ToDate"
    $labToDate.Size = New-Object System.Drawing.Size(80,40)
    $labToDate.Text = "To or less than"
     

    # ToDate Date Picker
    $ToDatePicker.DataBindings.DefaultDataSourceUpdateMode = 0
    $ToDatePicker.Location = New-Object System.Drawing.Point(100,330)
    $ToDatePicker.Name = "ToDatePicker"
    $ToDatePicker.Text = ""


    #Label Subject
    $labSubject.Location = New-Object System.Drawing.Point(5,370)
    $labSubject.Size = New-Object System.Drawing.Size(50,20)
    $labSubject.Name = "labSubject"
    $labSubject.Text = "Subject: "
     
     
    #TextBox Subject
    $txtBoxSubject.Location = New-Object System.Drawing.Point(100,370)
    $txtBoxSubject.Size = New-Object System.Drawing.Size(280,20)
    $txtBoxSubject.Name = "txtBoxSubject"
    $txtBoxSubject.Text = ""


    #Label FolderID
    $labFolderID.Location = New-Object System.Drawing.Point(5,400)
    $labFolderID.Size = New-Object System.Drawing.Size(55,20)
    $labFolderID.Name = "labFolderID"
    $labFolderID.Text = "FolderID:"
     
    #TextBox FolderID
    $txtBoxFolderID.Location = New-Object System.Drawing.Point(100,400)
    $txtBoxFolderID.Size = New-Object System.Drawing.Size(280,20)
    $txtBoxFolderID.Name = "txtBoxFolderID"
    $txtBoxFolderID.Text = ""


    #Adapting FolderID and TxtBoxFolderID based on the selection
    if($radiobutton4.Checked -or $radiobutton8.Checked -or $radiobutton9.Checked){
        $labFolderID.Location = New-Object System.Drawing.Point(5,285)
        $txtBoxFolderID.Location = New-Object System.Drawing.Point(100,285)
    }
    elseif($radiobutton12.Checked){
        $labFolderID.Size = New-Object System.Drawing.Size(95,20)
        $labFolderID.Text = "SourceFolderID:"
    } 
    elseif($radiobutton17.Checked){
        $labFolderID.Location = New-Object System.Drawing.Point(5,285)
        $labFolderID.Size = New-Object System.Drawing.Size(95,20)
        $labFolderID.Text = "E-mail Address:"
        $txtBoxFolderID.Location = New-Object System.Drawing.Point(100,285)
    }

    #Label Target FolderID
    $labTargetFolderID.Location = New-Object System.Drawing.Point(5,430)
    $labTargetFolderID.Size = New-Object System.Drawing.Size(95,20)
    $labTargetFolderID.Name = "labTargetFolderID"
    $labTargetFolderID.Text = "TargetFolderID:"
        
    #TextBox Target FolderID
    $txtBoxTargetFolderID.Location = New-Object System.Drawing.Point(100,430)
    $txtBoxTargetFolderID.Size = New-Object System.Drawing.Size(280,20)
    $txtBoxTargetFolderID.Name = "txtBoxTargetFolderID"
    $txtBoxTargetFolderID.Text = ""

    #Label Combobox
    $labelCombobox.Location = New-Object System.Drawing.Point(400,285)
    $labelCombobox.Size = New-Object System.Drawing.Size(80,35)
    $labelCombobox.Name = "labelCombobox"
    $labelCombobox.Text = "Delete Option"

    #ComboBox Menu
    $comboBoxMenu.DataBindings.DefaultDataSourceUpdateMode = 0
    $comboBoxMenu.FormattingEnabled = $True
    $comboBoxMenu.Location = New-Object System.Drawing.Point(485,285)
    $comboBoxMenu.Name = "comboBoxMenu"
    $comboBoxMenu.add_SelectedIndexChanged($handler_comboBoxMenu_SelectedIndexChanged)

    #Label ComboboxFolder
    $labelComboboxFolder.Location = New-Object System.Drawing.Point(5,285)
    $labelComboboxFolder.Size = New-Object System.Drawing.Size(50,35)
    $labelComboboxFolder.Name = "labelComboboxFolder"
    $labelComboboxFolder.Text = "Folder:"

    #ComboBoxFolder
    $comboBoxFolder.DataBindings.DefaultDataSourceUpdateMode = 0
    $comboBoxFolder.FormattingEnabled = $True
    $comboBoxFolder.Location = New-Object System.Drawing.Point(55,285)
    $comboBoxFolder.Size = New-Object System.Drawing.Size(70,35)
    $comboBoxFolder.Name = "comboBoxFolder"
    $comboBoxFolder.Items.Add("")|Out-Null
    $comboBoxFolder.Items.Add("Root")|Out-Null
    $comboBoxFolder.Items.Add("Calendar")|Out-Null
    $comboBoxFolder.Items.Add("Inbox")|Out-Null
    $comboBoxFolder.add_SelectedIndexChanged($handler_comboBoxFolder_SelectedIndexChanged)

    #Label ComboboxConfig
    $labelComboboxConfig.Location = New-Object System.Drawing.Point(145,285)
    $labelComboboxConfig.Size = New-Object System.Drawing.Size(75,35)
    $labelComboboxConfig.Name = "labelComboboxConfig"
    $labelComboboxConfig.Text = "Config Name:"

    #ComboBoxConfig
    $comboBoxConfig.DataBindings.DefaultDataSourceUpdateMode = 0
    $comboBoxConfig.FormattingEnabled = $True
    $comboBoxConfig.Location = New-Object System.Drawing.Point(225,285)
    $ComboboxConfig.Size = New-Object System.Drawing.Size(180,35)
    $comboBoxConfig.Name = "comboBoxConfig"
    $comboBoxConfig.Items.Add("")|Out-Null
    $comboBoxConfig.Items.Add("Aggregated.OwaUserConfiguration")|Out-Null
    $comboBoxConfig.Items.Add("UserConfigurationProperties.All")|Out-Null
    $comboBoxConfig.Items.Add("OWA.AttachmentDataProvider")|Out-Null
    $comboBoxConfig.Items.Add("OWA.AutocompleteCache")|Out-Null
    $comboBoxConfig.Items.Add("OWA.SessionInformation")|Out-Null
    $comboBoxConfig.Items.Add("OWA.UserOptions")|Out-Null
    $comboBoxConfig.Items.Add("OWA.ViewStateConfiguration")|Out-Null
    $comboBoxConfig.Items.Add("Suite.Storage")|Out-Null
    $comboBoxConfig.Items.Add("UM.E14.PersonalAutoAttendants")|Out-Null
    $comboBoxConfig.Items.Add("CleanFinders")|Out-Null
    $comboBoxConfig.add_SelectedIndexChanged($handler_comboBoxConfig_SelectedIndexChanged)
        
    if($radiobutton4.Checked){
        $PremiseForm.Controls.Add($labFolderID)
        $PremiseForm.Controls.Add($txtBoxFolderID)
        }
    elseif($radiobutton7.Checked){
        $PremiseForm.Controls.Add($labFolderID)
        $PremiseForm.Controls.Add($txtBoxFolderID)
        $PremiseForm.Controls.Add($labFromDate)
        $PremiseForm.Controls.Add($FromDatePicker)
        $PremiseForm.Controls.Add($labToDate)
        $PremiseForm.Controls.Add($ToDatePicker) 
        $PremiseForm.Controls.Add($labSubject)
        $PremiseForm.Controls.Add($txtBoxSubject)
    }
    elseif($radiobutton8.Checked){
        $labFolderID.Size = New-Object System.Drawing.Size(95,20)
        $labFolderID.Text = "Folder Name:"
        $PremiseForm.Controls.Add($labFolderID)
        $PremiseForm.Controls.Add($txtBoxFolderID)
    }
    elseif($radiobutton9.Checked){
        $comboBoxMenu.Items.Add("")|Out-Null
        $comboBoxMenu.Items.Add("HardDelete")|Out-Null
        $comboBoxMenu.Items.Add("MoveToDeletedItems")|Out-Null
        $comboBoxMenu.SelectedItem = "MoveToDeletedItems"
        $PremiseForm.Controls.Add($labFolderID)
        $PremiseForm.Controls.Add($txtBoxFolderID)
        $PremiseForm.Controls.Add($labelCombobox)
        $PremiseForm.Controls.Add($comboBoxMenu)
    }
    elseif($radiobutton12.Checked){
        $PremiseForm.Controls.Add($labFromDate)
        $PremiseForm.Controls.Add($FromDatePicker)
        $PremiseForm.Controls.Add($labToDate)
        $PremiseForm.Controls.Add($ToDatePicker) 
        $PremiseForm.Controls.Add($labSubject)
        $PremiseForm.Controls.Add($txtBoxSubject)
        $PremiseForm.Controls.Add($labFolderID)
        $PremiseForm.Controls.Add($txtBoxFolderID)
        $PremiseForm.Controls.Add($labTargetFolderID)
        $PremiseForm.Controls.Add($txtBoxTargetFolderID)
    }
    elseif($radiobutton13.Checked){
        $comboBoxMenu.Items.Add("")|Out-Null
        $comboBoxMenu.Items.Add("SoftDelete")|Out-Null
        $comboBoxMenu.Items.Add("HardDelete")|Out-Null
        $comboBoxMenu.Items.Add("MoveToDeletedItems")|Out-Null
        $comboBoxMenu.SelectedItem = "MoveToDeletedItems"
        $PremiseForm.Controls.Add($labFromDate)
        $PremiseForm.Controls.Add($FromDatePicker)
        $PremiseForm.Controls.Add($labToDate)
        $PremiseForm.Controls.Add($ToDatePicker) 
        $PremiseForm.Controls.Add($labSubject)
        $PremiseForm.Controls.Add($txtBoxSubject)
        $PremiseForm.Controls.Add($labFolderID)
        $PremiseForm.Controls.Add($txtBoxFolderID)
        $PremiseForm.Controls.Add($labelCombobox)
        $PremiseForm.Controls.Add($comboBoxMenu)
    }
    elseif($radiobutton15.Checked){
        $comboBoxMenu.Items.Add("")|Out-Null
        $comboBoxMenu.Items.Add("Normal")|Out-Null
        $comboBoxMenu.Items.Add("Personal")|Out-Null
        $comboBoxMenu.Items.Add("Private")|Out-Null
        $comboBoxMenu.Items.Add("Confidential")|Out-Null
        $comboBoxMenu.SelectedItem = "Normal"
        $PremiseForm.Controls.Add($labFromDate)
        $PremiseForm.Controls.Add($FromDatePicker)
        $PremiseForm.Controls.Add($labToDate)
        $PremiseForm.Controls.Add($ToDatePicker) 
        $PremiseForm.Controls.Add($labSubject)
        $PremiseForm.Controls.Add($txtBoxSubject)
        $PremiseForm.Controls.Add($labFolderID)
        $PremiseForm.Controls.Add($txtBoxFolderID)
        $PremiseForm.Controls.Add($labelCombobox)
        $PremiseForm.Controls.Add($comboBoxMenu)
    }
    elseif($radiobutton16.Checked){
        $PremiseForm.Controls.Add($labelComboboxFolder)
        $PremiseForm.Controls.Add($comboBoxFolder)
        $PremiseForm.Controls.Add($labelComboboxConfig)
        $PremiseForm.Controls.Add($comboBoxConfig)
    }
    elseif($radiobutton17.Checked){
        $PremiseForm.Controls.Add($labFolderID)
        $PremiseForm.Controls.Add($txtBoxFolderID)
    }
    $PremiseForm.refresh()

    }

    $handler_comboBoxMenu_SelectedIndexChanged= {
    # Get the Event ID when item is selected
        $Global:ComboOption = $comboBoxMenu.selectedItem.ToString()
    }

    $handler_comboBoxFolder_SelectedIndexChanged= {
    # Get the Event ID when item is selected
        $Global:ComboOption1 = $comboBoxFolder.selectedItem.ToString()
    }

    $handler_comboBoxConfig_SelectedIndexChanged= {
        # Get the Event ID when item is selected
            $Global:ComboOption2 = $comboBoxConfig.selectedItem.ToString()
    }

    $handler_labImpersonationHelp_Click={
        [Microsoft.VisualBasic.Interaction]::MsgBox("In order to use Impersonation, we must first assign proper ManagementRole to the 'administrative' account that run the different options.
    New-ManagementRoleAssignment –Name:impersonationAssignmentName –Role:ApplicationImpersonation –User:<Account>

    More info at: https://msdn.microsoft.com/en-us/library/bb204095(exchg.140).aspx

    Press CTRL + C to copy this message to clipboard.",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
    }

    $OnLoadMainWindow_StateCorrection={#Correct the initial state of the form to prevent the .Net maximized form issue
        $PremiseForm.WindowState = $InitialFormWindowState
    }


    #----------------------------------------------
    #region Generated Form Code


    $PremiseForm.Controls.Add($radiobutton1)
    $PremiseForm.Controls.Add($radiobutton2)
    $PremiseForm.Controls.Add($radiobutton3)
    $PremiseForm.Controls.Add($radiobutton4)
    $PremiseForm.Controls.Add($radiobutton5)
    $PremiseForm.Controls.Add($radiobutton6)
    $PremiseForm.Controls.Add($radiobutton7)
    $PremiseForm.Controls.Add($radiobutton8)
    $PremiseForm.Controls.Add($radiobutton9)
    $PremiseForm.Controls.Add($radiobutton10)
    $PremiseForm.Controls.Add($radiobutton11)
    $PremiseForm.Controls.Add($radiobutton12)
    $PremiseForm.Controls.Add($radiobutton13)
    $PremiseForm.Controls.Add($radiobutton14)
    $PremiseForm.Controls.Add($radiobutton15)
    $PremiseForm.Controls.Add($radiobutton16)
    $PremiseForm.Controls.Add($radiobutton17)
    $statusBar = New-Object System.Windows.Forms.StatusBar
    $statusBar.Name = "statusBar"
    $statusBar.Text = "Ready..."
    $PremiseForm.Controls.Add($statusBar)
    $PremiseForm.ClientSize = New-Object System.Drawing.Size(800,720)
    $PremiseForm.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $PremiseForm.Name = "form1"
    $PremiseForm.Text = "Managing user: " + $email + ". Choose your Option"
    $PremiseForm.StartPosition = "CenterScreen"
    $PremiseForm.KeyPreview = $True
    $PremiseForm.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$PremiseForm.Close()} })
    #
    # radiobutton1
    #
    $radiobutton1.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton1.Location = New-Object System.Drawing.Point(20,20)
    $radiobutton1.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton1.TabIndex = 1
    $radiobutton1.Text = "1 - List Folders in Root"
    $radioButton1.Checked = $true
    $radiobutton1.UseVisualStyleBackColor = $True
    $radiobutton1.Add_Click({& $ExpandFilters})
    #
    # radiobutton2
    #
    $radiobutton2.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton2.Location = New-Object System.Drawing.Point(20,50)
    $radiobutton2.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton2.TabIndex = 2
    $radiobutton2.Text = "2 - List Folders in Archive Root"
    $radioButton2.Checked = $false
    $radiobutton2.UseVisualStyleBackColor = $True
    $radiobutton2.Add_Click({& $ExpandFilters})
    #
    # radiobutton3
    #
    $radiobutton3.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton3.Location = New-Object System.Drawing.Point(20,80)
    $radiobutton3.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton3.TabIndex = 3
    $radiobutton3.Text = "3 - List Folders in Public Folder Root"
    $radiobutton3.Checked = $false
    $radiobutton3.UseVisualStyleBackColor = $True
    $radiobutton3.Add_Click({& $ExpandFilters})
    #
    # radiobutton4
    #
    $radiobutton4.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton4.Location = New-Object System.Drawing.Point(20,110)
    $radiobutton4.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton4.TabIndex = 4
    $radiobutton4.Text = "4 - List subFolders from a desired Parent Folder"
    $radiobutton4.Checked = $false
    $radiobutton4.UseVisualStyleBackColor = $True
    $radiobutton4.Add_Click({& $ExpandFilters})
    #
    # radiobutton5
    #
    $radiobutton5.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton5.Location = New-Object System.Drawing.Point(20,140)
    $radiobutton5.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton5.Tabindex = 5
    $radiobutton5.Text = "5 - List folders in Recoverable Items Root folder"
    $radiobutton5.Checked = $false
    $radiobutton5.UseVisualStyleBackColor = $True
    $radiobutton5.Add_Click({& $ExpandFilters})
    #
    # radiobutton6
    #
    $radiobutton6.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton6.Location = New-Object System.Drawing.Point(20,170)
    $radiobutton6.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton6.TabIndex = 6
    $radiobutton6.Text = "6 - List folders in Recoverable Items folder in Archive"
    $radiobutton6.Checked = $false
    $radiobutton6.UseVisualStyleBackColor = $True
    $radiobutton6.Add_Click({& $ExpandFilters})
    #
    # radiobutton7
    #
    $radiobutton7.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton7.Location = New-Object System.Drawing.Point(20,200)
    $radiobutton7.Name = "radiobutton7"
    $radiobutton7.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton7.TabIndex = 7
    $radiobutton7.Text = "7 - List Items in a desired Folder"
    $radiobutton7.Checked = $false
    $radiobutton7.UseVisualStyleBackColor = $True
    $radiobutton7.Add_Click({& $ExpandFilters})
    #
    # radiobutton8
    #
    $radiobutton8.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton8.Location = New-Object System.Drawing.Point(20,230)
    $radiobutton8.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton8.TabIndex = 8
    $radiobutton8.Text = "8 - Create a custom Folder in Root"
    $radiobutton8.Checked = $false
    $radiobutton8.UseVisualStyleBackColor = $True
    $radiobutton8.Add_Click({& $ExpandFilters})
    #
    # radiobutton9
    #
    $radiobutton9.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton9.Location = New-Object System.Drawing.Point(20,260)
    $radiobutton9.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton9.TabIndex = 9
    $radiobutton9.Text = "9 - Delete a Folder"
    $radiobutton9.Checked = $false
    $radiobutton9.UseVisualStyleBackColor = $True
    $radiobutton9.Add_Click({& $ExpandFilters})
    #
    # radiobutton10
    #
    $radiobutton10.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton10.Location = New-Object System.Drawing.Point(400,20)
    $radiobutton10.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton10.TabIndex = 10
    $radiobutton10.Text = "10 - Get user's Inbox Rules"
    $radiobutton10.Checked = $false
    $radiobutton10.UseVisualStyleBackColor = $True
    $radiobutton10.Add_Click({& $ExpandFilters})
    #
    # radiobutton11
    #
    $radiobutton11.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton11.Location = New-Object System.Drawing.Point(400,50)
    $radiobutton11.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton11.TabIndex = 11
    $radiobutton11.Text = "11 - Get user's OOF Settings"
    $radiobutton11.Checked = $false
    $radiobutton11.UseVisualStyleBackColor = $True
    $radiobutton11.Add_Click({& $ExpandFilters})
    #
    # radiobutton12
    #
    $radiobutton12.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton12.Location = New-Object System.Drawing.Point(400,80)
    $radiobutton12.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton12.TabIndex = 12
    $radiobutton12.Text = "12 - Move items between folders"
    $radiobutton12.Checked = $false
    $radiobutton12.UseVisualStyleBackColor = $True
    $radiobutton12.Add_Click({& $ExpandFilters})
    #
    # radiobutton13
    #
    $radiobutton13.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton13.Location = New-Object System.Drawing.Point(400,110)
    $radiobutton13.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton13.TabIndex = 13
    $radiobutton13.Text = "13 - Delete a subset of items in a folder"
    $radiobutton13.Checked = $false
    $radiobutton13.UseVisualStyleBackColor = $True
    $radiobutton13.Add_Click({& $ExpandFilters})
    #
    # radiobutton14
    #
    $radiobutton14.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton14.Location = New-Object System.Drawing.Point(400,140)
    $radiobutton14.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton14.TabIndex = 14
    $radiobutton14.Text = "14 - Get user's Delegate information"
    $radiobutton14.Checked = $false
    $radiobutton14.UseVisualStyleBackColor = $True
    $radiobutton14.Add_Click({& $ExpandFilters})
    #
    # radiobutton15
    #
    $radiobutton15.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton15.Location = New-Object System.Drawing.Point(400,170)
    $radiobutton15.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton15.TabIndex = 15
    $radiobutton15.Text = "15 - Change sensitivity to items in a folder"
    $radiobutton15.Checked = $false
    $radiobutton15.UseVisualStyleBackColor = $True
    $radiobutton15.Add_Click({& $ExpandFilters})
    #
    # radiobutton16
    #
    $radiobutton16.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton16.Location = New-Object System.Drawing.Point(400,200)
    $radiobutton16.Size = New-Object System.Drawing.Size(300,15)
    $radiobutton16.TabIndex = 16
    $radiobutton16.Text = "16 - Remove OWA configurations"
    $radiobutton16.Checked = $false
    $radiobutton16.UseVisualStyleBackColor = $True
    $radiobutton16.Add_Click({& $ExpandFilters})
    #
    # radiobutton17
    #
    $radiobutton17.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton17.Location = New-Object System.Drawing.Point(400,230)
    $radiobutton17.Size = New-Object System.Drawing.Size(175,15)
    $radiobutton17.TabIndex = 17
    $radiobutton17.Text = "17 - Switch to another Mailbox"
    $radiobutton17.Checked = $false
    $radiobutton17.UseVisualStyleBackColor = $True
    $radiobutton17.Add_Click({& $ExpandFilters})
    #
    # Label Impersonation Help
    #
    $labImpersonationHelp.Location = New-Object System.Drawing.Point(380,231)
    $labImpersonationHelp.Size = New-Object System.Drawing.Size(10,20)
    $labImpersonationHelp.Name = "labImpersonation"
    $labImpersonationHelp.ForeColor = "Blue"
    $labImpersonationHelp.Text = "?"
    $labImpersonationHelp.add_Click($handler_labImpersonationHelp_Click)
    $PremiseForm.Controls.Add($labImpersonationHelp)


    #"Go" button
    $Global:buttonGo.Dispose()

    $Global:buttonGo2 = New-Object System.Windows.Forms.Button
    $buttonGo2.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGo2.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
    $buttonGo2.Location = New-Object System.Drawing.Point(700,20)
    $buttonGo2.Size = New-Object System.Drawing.Size(50,25)
    $buttonGo2.TabIndex = 17
    $buttonGo2.Name = "Go"
    $buttonGo2.Text = "Go"
    $buttonGo2.UseVisualStyleBackColor = $True
    $buttonGo2.add_Click({
            if($radiobutton1.Checked){& $Process1to6}
            elseif($radiobutton2.Checked){& $Process1to6}
            elseif($radiobutton3.Checked){& $Process1to6}
            elseif($radiobutton4.Checked){& $Process1to6}
            elseif($radiobutton5.Checked){& $Process1to6}
            elseif($radiobutton6.Checked){& $Process1to6}
            elseif($radiobutton7.Checked){& $process7}
            elseif($radiobutton8.Checked){& $process8}
            elseif($radiobutton9.Checked){& $process9}
            elseif($radiobutton10.Checked){& $process10}
            elseif($radiobutton11.Checked){& $process11}
            elseif($radiobutton12.Checked){& $process12}
            elseif($radiobutton13.Checked){& $process13}
            elseif($radiobutton14.Checked){& $process14}
            elseif($radiobutton15.Checked){& $process15}
            elseif($radiobutton16.Checked){& $Process16}
            elseif($radiobutton17.Checked){& $Process17}
    })
    $PremiseForm.Controls.Add($buttonGo2)


    #"Exit" button
    $buttonExit.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonExit.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
    $buttonExit.Location = New-Object System.Drawing.Point(700,50)
    $buttonExit.Size = New-Object System.Drawing.Size(50,25)
    $buttonExit.TabIndex = 17
    $buttonExit.Name = "Exit"
    $buttonExit.Text = "Exit"
    $buttonExit.UseVisualStyleBackColor = $True
    $PremiseForm.Controls.Add($buttonExit)

    #TextBox results
    $txtBoxResults.DataBindings.DefaultDataSourceUpdateMode = 0
    $txtBoxResults.Location = New-Object System.Drawing.Point(5,460)
    $txtBoxResults.Size = New-Object System.Drawing.Size(790,240)
    $txtBoxResults.Name = "TextResults"
    $txtBoxResults.BackColor = [System.Drawing.Color]::White
    $txtBoxResults.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $Font = New-Object System.Drawing.Font("Consolas",8)
    $txtBoxResults.Font = $Font 
    $PremiseForm.Controls.Add($txtBoxResults)


    #dataGrid

    $dgResults.Anchor = 15
    $dgResults.DataBindings.DefaultDataSourceUpdateMode = 0
    $dgResults.DataMember = ""
    $dgResults.Location = New-Object System.Drawing.Point(5,460)
    $dgResults.Size = New-Object System.Drawing.Size(790,240)
    $dgResults.Name = "dgResults"
    $dgResults.ReadOnly = $True
    $dgResults.RowHeadersVisible = $False
    $dgResults.Visible = $False
    $dgResults.AllowUserToOrderColumns = $True
    $dgResults.AllowUserToResizeColumns = $True
    $PremiseForm.Controls.Add($dgResults)


    #endregion Generated Form Code

    # Show Form
    #Save the initial state of the form
    $InitialFormWindowState = $PremiseForm.WindowState
    #Init the OnLoad event to correct the initial state of the form
    $PremiseForm.add_Load($OnLoadMainWindow_StateCorrection)
    $PremiseForm.Add_Shown({$PremiseForm.Activate()})
    $PremiseForm.ShowDialog()| Out-Null
    #exit if 'Exit' button is pushed
    if($buttonExit.IsDisposed){return} 

    } #End Function
     
    #Call the Function
    GenerateForm
}