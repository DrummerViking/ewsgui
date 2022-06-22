Function Connect-EWSService {
    <#
    .SYNOPSIS
    Function to Create service object and authenticate the user.
    
    .DESCRIPTION
    This function will create the service object.
    Will opt to the user to select connection either to On-premises or Exchange Online.
    Will use basic auth to connect to on-premises. Endpoint will be discovered using Autodiscover.
    Will use modern auth to connect to Exchange Online. Endpoint is hard-coded to EXO EWS URL.
    
    .PARAMETER ClientID
    String parameter with the ClientID (or AppId) of your AzureAD Registered App.

    .PARAMETER TenantID
    String parameter with the TenantID your AzureAD tenant.

    .PARAMETER ClientSecret
    String parameter with the Client Secret which is configured in the AzureAD App.

    .EXAMPLE
    PS C:\> Connect-EWSService
    Creates service object and authenticate the user.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidGlobalVars", "")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseDeclaredVarsMoreThanAssignments", "")]
    [Cmdletbinding()]
    param(
        [String] $ClientID,

        [String] $TenantID,

        [String] $ClientSecret
    )
    # Choosing if connection is to Office 365 or an Exchange on-premises
    $PremiseForm.Controls.Add($radiobutton1)
    $PremiseForm.Controls.Add($radiobutton2)
    $PremiseForm.Controls.Add($radiobutton3)
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
        if ( -not(Get-Module Microsoft.Identity.Client -ListAvailable) -and -not(Get-Module Microsoft.Identity.Client) ) {
            Install-Module Microsoft.Identity.Client -Force -ErrorAction Stop
        }
        Import-Module Microsoft.Identity.Client

        # Connecting using Oauth with Application permissions
        if ( -not[String]::IsNullOrEmpty($ClientID) -or -not[String]::IsNullOrEmpty($TenantID) -or -not[String]::IsNullOrEmpty($ClientSecret) ) {
            $cid = $ClientID
            $tid = $TenantID
            $cs = $clientSecret

            $ccaOptions = [Microsoft.Identity.Client.ConfidentialClientApplicationOptions]::new()
            $ccaOptions.ClientID = $cid
            $ccaOptions.TenantID = $Tid
            $ccaOptions.ClientSecret = $cs
            $ccaBuilder = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::CreateWithApplicationOptions($ccaOptions)
            $cca = $ccaBuilder.Build()
            $scopes = New-Object System.Collections.Generic.List[string]
            $scopes.Add("https://outlook.office365.com/.default")
            $authResult = $cca.AcquireTokenForClient($scopes)
            $token = $authResult.ExecuteAsync()
            while ( $token.IsCompleted -eq $False ) { <# Waiting for token auth flow to complete #> }
            if ($token.Status -eq "Faulted" -and $token.Exception.Message.StartsWith("One or more errors occurred. (ActiveX control '8856f961-340a-11d0-a96b-00c04fd705a2'")) {
                Write-PSFHostColor -String "Known issue occurred. There is work in progress to fix authentication flow." -DefaultColor Red
                Write-PSFHostColor -String "Failed to obtain authentication token. Exiting script. Please rerun the script again and it should work." -DefaultColor Red
                exit
            }
            Write-PSFMessage -Level Important -Message "Connected using Application permissions with passed ClientID, TenantID and ClientSecret"
        }
        elseif (
            $null -ne (Get-pSFConfig -Module EwsGui -Name ClientID).value -and `
            $null -ne (Get-pSFConfig -Module EwsGui -Name TenantID).value -and `
            $null -ne (Get-pSFConfig -Module EwsGui -Name ClientSecret).value
        ) {
            $cid = (Get-pSFConfig -Module EwsGui -Name ClientID).value
            $tid = (Get-pSFConfig -Module EwsGui -Name TenantID).value
            $cs = (Get-pSFConfig -Module EwsGui -Name ClientSecret).value

            $ccaOptions = [Microsoft.Identity.Client.ConfidentialClientApplicationOptions]::new()
            $ccaOptions.ClientId = $cid
            $ccaOptions.TenantID = $Tid
            $ccaOptions.ClientSecret = $cs
            $ccaBuilder = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::CreateWithApplicationOptions($ccaOptions)
            $cca = $ccaBuilder.Build()
            $scopes = New-Object System.Collections.Generic.List[string]
            $scopes.Add("https://outlook.office365.com/.default")
            $authResult = $cca.AcquireTokenForClient($scopes)
            $token = $authResult.ExecuteAsync()
            while ( $token.IsCompleted -eq $False ) { <# Waiting for token auth flow to complete #> }
            if ($token.Status -eq "Faulted" -and $token.Exception.Message.StartsWith("One or more errors occurred. (ActiveX control '8856f961-340a-11d0-a96b-00c04fd705a2'")) {
                Write-PSFHostColor -String "Known issue occurred. There is work in progress to fix authentication flow." -DefaultColor Red
                Write-PSFHostColor -String "Failed to obtain authentication token. Exiting script. Please rerun the script again and it should work." -DefaultColor Red
                exit
            }
            Write-PSFMessage -Level Important -Message "Connected using Application permissions with registered ClientID, TenantID and ClientSecret embedded to the module."
        }
        else {
            # Connecting using Oauth with delegated permissions
            $pcaOptions = [Microsoft.Identity.Client.PublicClientApplicationOptions]::new()
            $pcaOptions.ClientId = "8799ab60-ace5-4bda-b31f-621c9f6668db"
            $pcaOptions.RedirectUri = "http://localhost/code"
            $pcaBuilder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::CreateWithApplicationOptions($pcaOptions)
            $pca = $pcaBuilder.Build()
            $scopes = New-Object System.Collections.Generic.List[string]
            $scopes.Add("https://outlook.office365.com/.default")
            #$scopes.Add("https://outlook.office.com/EWS.AccessAsUser.All")
            $authResult = $pca.AcquireTokenInteractive($scopes)
            $global:token = $authResult.ExecuteAsync()
            while ( $token.IsCompleted -eq $False ) { <# Waiting for token auth flow to complete #> }
            if ($token.Status -eq "Faulted" -and $token.Exception.Message.StartsWith("One or more errors occurred. (ActiveX control '8856f961-340a-11d0-a96b-00c04fd705a2'")) {
                Write-PSFHostColor -String "Known issue occurred. There is work in progress to fix authentication flow." -DefaultColor Red
                Write-PSFHostColor -String "Failed to obtain authentication token. Exiting script. Please rerun the script again and it should work." -DefaultColor Red
                exit
            }
            Write-PSFMessage -Level Important -Message "Connected using Delegated permissions with: $($token.result.Account.Username)"
        }
        
        $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($Token.Result.AccessToken)
        $Global:email = $Token.Result.Account.Username
        $service.Url = New-Object Uri("https://outlook.office365.com/ews/exchange.asmx")
    }
    else {
        $psCred = Get-Credential -Message "Type your credentials or Administrator credentials"
        $Global:email = $psCred.UserName
        $exchangeCredentials = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(), $psCred.GetNetworkCredential().password.ToString())
        # setting Autodiscover endpoint
        $service.EnableScpLookup = $True
        $service.AutodiscoverUrl($email, { $true })
    }
    $Service.Credentials = $exchangeCredentials

    return $service
}