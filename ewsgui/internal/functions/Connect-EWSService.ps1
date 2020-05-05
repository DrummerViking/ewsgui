Function Connect-EWSService {
    <#
    .SYNOPSIS
    Function to Create service object and authenticate the user.
    
    .DESCRIPTION
    This function will create the service object.
    Will opt to the user to select connection either to On-premises or Exchange Online.
        Will use basic auth to connect to on-premises. Endpoint will be discovered using Autodiscover.
        Will use modern auth to connect to Exchange Online. Endpoint is hard-coded to EXO EWS URL.
    
    .EXAMPLE
    PS C:\> Connect-EWSService
    Creates service object and authenticate the user.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidGlobalVars", "")]
    [Cmdletbinding()]
    param(
        # Parameters
    )
    # Choosing if connection is to Office 365 or an Exchange on-premises
    $PremiseForm.Controls.Add($radiobutton1)
    $PremiseForm.Controls.Add($radiobutton2)
    $PremiseForm.Controls.Add($radiobutton3)
    #$PremiseForm.Controls.Add($radiobutton4)
    $PremiseForm.ClientSize = New-Object System.Drawing.Size(250, 160)
    $PremiseForm.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
    $PremiseForm.Name = "form1"
    $PremiseForm.Text = "Choose your Exchange version"
    #
    # radiobutton1
    #
    $radiobutton1.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
    $radiobutton1.Location = New-Object System.Drawing.Point(20, 20)
    $radiobutton1.Size = New-Object System.Drawing.Size(150, 25)
    $radiobutton1.TabStop = $True
    $radiobutton1.Text = "Exchange 2010"
    $radioButton1.Checked = $true
    $radiobutton1.UseVisualStyleBackColor = $True
    #
    # radiobutton2
    #
    $radiobutton2.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
    $radiobutton2.Location = New-Object System.Drawing.Point(20, 55)
    $radiobutton2.Size = New-Object System.Drawing.Size(150, 30)
    $radiobutton2.TabStop = $True
    $radiobutton2.Text = "Exchange 2013/2016/2019"
    $radioButton2.Checked = $false
    $radiobutton2.UseVisualStyleBackColor = $True
    #
    # radiobutton3
    #
    $radiobutton3.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
    $radiobutton3.Location = New-Object System.Drawing.Point(20, 95)
    $radiobutton3.Size = New-Object System.Drawing.Size(150, 25)
    $radiobutton3.TabStop = $True
    $radiobutton3.Text = "Office365"
    $radiobutton3.Checked = $false
    $radiobutton3.UseVisualStyleBackColor = $True
    #
    # radiobutton4
    #
    #$radiobutton4.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
    #$radiobutton4.Location = New-Object System.Drawing.Point(20, 110)
    #$radiobutton4.Size = New-Object System.Drawing.Size(150, 30)
    #$radiobutton4.Text = "Office365"
    #$radiobutton4.Checked = $false
    #$radiobutton4.UseVisualStyleBackColor = $True

    #"Go" button
    $buttonGo.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGo.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 170
    $System_Drawing_Point.Y = 20
    $buttonGo.Location = $System_Drawing_Point
    $buttonGo.Name = "Go"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 25
    $System_Drawing_Size.Width = 50
    $buttonGo.Size = $System_Drawing_Size
    $buttonGo.Text = "Go"
    $buttonGo.UseVisualStyleBackColor = $True
    $buttonGo.add_Click( {
            if ($radiobutton1.Checked) { $Global:option = "Exchange2010_SP2" }
            elseif ($radiobutton2.Checked) { $Global:option = "Exchange2013_SP1" }
            elseif ($radiobutton3.Checked) { $Global:option = "Exchange2013_SP1" }
            $PremiseForm.Hide()
        })
    $PremiseForm.Controls.Add($buttonGo)

    #"Exit" button
    $buttonExit.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonExit.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 170
    $System_Drawing_Point.Y = 50
    $buttonExit.Location = $System_Drawing_Point
    $buttonExit.Name = "Exit"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 25
    $System_Drawing_Size.Width = 50
    $buttonExit.Size = $System_Drawing_Size
    $buttonExit.Text = "Exit"
    $buttonExit.UseVisualStyleBackColor = $True
    $buttonExit.add_Click( { $PremiseForm.Close() ; $buttonExit.Dispose() })
    $PremiseForm.Controls.Add($buttonExit)

    #Show Form
    $PremiseForm.Add_Shown( { $PremiseForm.Activate() })
    $PremiseForm.ShowDialog() | Out-Null
    #exit if 'Exit' button is pushed
    if ($buttonExit.IsDisposed) { return }

    #creating service object
    $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$option
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
 
    if ($radiobutton3.Checked) {
        #Getting oauth credentials
        $Folderpath = (Get-Module azuread -ListAvailable | Sort-Object Version -Descending)[0].Path
        $path = join-path (split-path $Folderpath -parent) 'Microsoft.IdentityModel.Clients.ActiveDirectory.dll'
        Add-Type -Path $path

        $authenticationContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.windows.net/common", $False)
        $platformParameters = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters([Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Always)
        $resourceUri = "https://outlook.office365.com"
        $AppId = "8799ab60-ace5-4bda-b31f-621c9f6668db"
        $redirectUri = New-Object Uri("http://localhost/code")

        Write-PSFMessage -Level Verbose -FunctionName "EWSGui" -Message "Looking in token cache"
        $authenticationResult = $authenticationContext.AcquireTokenSilentAsync($resourceUri, $AppId)

        while ($authenticationResult.IsCompleted -ne $true) { Start-Sleep -Milliseconds 500; Write-PSFMessage -Level Verbose -FunctionName "EWSGui" -Message "sleep" }

        # Check if we failed to get the token
        if (!($authenticationResult.IsFaulted -eq $false)) {

            Write-PSFMessage -Level Warning -Message "Acquire token silent failed"
            switch ($authenticationResult.Exception.InnerException.ErrorCode) {
                failed_to_acquire_token_silently {
                    # do nothing since we pretty much expect this to fail
                    Write-PSFMessage -Level Verbose -FunctionName "EWSGui" -Message "Cache miss, asking for credentials"
                    $authenticationResult = $authenticationContext.AcquireTokenAsync($resourceUri, $AppId, $redirectUri, $platformParameters)

                    while ($authenticationResult.IsCompleted -ne $true) { Start-Sleep -Milliseconds 500; Write-PSFMessage -Level Verbose -FunctionName "EWSGui" -Message "sleep" }
                }
                multiple_matching_tokens_detected {
                    # we could clear the cache here since we don't have a UPN, but we are just going to move on to prompting
                    Write-PSFMessage -Level Verbose -FunctionName "EWSGui" -Message "Multiple matching entries found, asking for credentials"
                    $authenticationResult = $authenticationContext.AcquireTokenAsync($resourceUri, $AppId, $redirectUri, $platformParameters)

                    while ($authenticationResult.IsCompleted -ne $true) { Start-Sleep -Milliseconds 500; Write-PSFMessage -Level Verbose -FunctionName "EWSGui" -Message "sleep" }
                }
                Default { Write-PSFMessage -Level Warning -Message "Unknown Token Error $authenticationResult.Exception.InnerException.ErrorCode" -ErrorRecord $_ }
            }
        }
        $Global:email = $authenticationResult.Result.UserInfo.DisplayableId
        $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($authenticationResult.Result.AccessToken)
        $service.Url = New-Object Uri("https://outlook.office365.com/ews/exchange.asmx")
    }
    else {
        $psCred = Get-Credential -Message "Type your credentials or Administrator credentials"
        $Global:email = $psCred.UserName
        $exchangeCredentials = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())
        # setting Autodiscover endpoint
        $service.EnableScpLookup = $True
        $service.AutodiscoverUrl($email,{$true})
    }
    $Service.Credentials = $exchangeCredentials

    return $service
}