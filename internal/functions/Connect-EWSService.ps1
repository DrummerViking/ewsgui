Function Connect-EWSService {
    <#
    .SYNOPSIS
    Short description
    
    .DESCRIPTION
    Long description
    
    .EXAMPLE
    An example
    
    #>

    # Choosing if connection is to Office 365 or an Exchange on-premises
    $PremiseForm.Controls.Add($radiobutton1)
    $PremiseForm.Controls.Add($radiobutton2)
    $PremiseForm.Controls.Add($radiobutton3)
    $PremiseForm.Controls.Add($radiobutton4)
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
    $radiobutton1.Text = "Exchange 2007"
    $radioButton1.Checked = $true
    $radiobutton1.UseVisualStyleBackColor = $True
    #
    # radiobutton2
    #
    $radiobutton2.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton2.Location = New-Object System.Drawing.Point(20, 50)
    $radiobutton2.Size = New-Object System.Drawing.Size(150, 20)
    $radiobutton2.TabStop = $True
    $radiobutton2.Text = "Exchange 2010"
    $radioButton2.Checked = $false
    $radiobutton2.UseVisualStyleBackColor = $True
    #
    # radiobutton3
    #
    $radiobutton3.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton3.Location = New-Object System.Drawing.Point(20, 80)
    $radiobutton3.Size = New-Object System.Drawing.Size(150, 25)
    $radiobutton3.TabStop = $True
    $radiobutton3.Text = "Exchange 2013/2016"
    $radiobutton3.Checked = $false
    $radiobutton3.UseVisualStyleBackColor = $True
    #
    # radiobutton4
    #
    $radiobutton4.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton4.Location = New-Object System.Drawing.Point(20, 110)
    $radiobutton4.Size = New-Object System.Drawing.Size(150, 30)
    $radiobutton4.Text = "Office365"
    $radiobutton4.Checked = $false
    $radiobutton4.UseVisualStyleBackColor = $True

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
            if ($radiobutton1.Checked) { $Global:option = "Exchange2007_SP1" }
            elseif ($radiobutton2.Checked) { $Global:option = "Exchange2010_SP2" }
            elseif ($radiobutton3.Checked) { $Global:option = "Exchange2013_SP1" }
            elseif ($radiobutton4.Checked) { $Global:option = "Exchange2013_SP1" }
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
 
    #setting credentials
    $psCred = Get-Credential -Message "Type your credentials or Administrator credentials"
    $Global:email = $psCred.UserName
    $authenticationContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/$TenantId", $False)
    $platformParameters = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters([Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Always)
    $redirectUri = New-Object Uri("http://localhost/code")
    $AADCredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential" -ArgumentList $psCred.UserName, $psCred.Password
    $authenticationResult = $authenticationContext.AcquireTokenAsync("https://outlook.office365.com", $AppDetails.AppId, $AADCredential)
    if ($authenticationResult.Exception.InnerException.ErrorCode -eq 'interaction_required' ) {
        $authenticationResult = $authenticationContext.AcquireTokenAsync("https://outlook.office365.com", $AppDetails.AppId, $redirectUri, $platformParameters)
    }
    $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($authenticationResult.Result.AccessToken)
    $Service.Credentials =  $exchangeCredentials
    $service.Url = New-Object Uri("https://outlook.office365.com/ews/exchange.asmx")

    return $service
}