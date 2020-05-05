Function Import-EWSDLL {
    <#
    .SYNOPSIS
    Loads EWS DLL into the session
    
    .DESCRIPTION
    Search and Loads EWS DLL into the session
    
    .EXAMPLE
    PS C:\> Import-EWSDLL
    Search and Loads EWS DLL into the session
    
    #>
    #[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseApprovedVerbs", "")]
    [Cmdletbinding()]
    param(
        # Parameters
    )
    Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] This script requires at least EWS API 2.1" -DefaultColor Yellow
    
    # Locating DLL location either in working path, in EWS API 2.1 path or in EWS API 2.2 path
    $Directory = ".\"
    $EWS = Join-Path $Directory "Microsoft.Exchange.WebServices.dll"
    $test = Test-Path -Path $EWS
    if ($test -eq $False) {
        Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] EWS DLL in local path not found" -DefaultColor Cyan
        $test2 = Test-Path -Path "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.*\Microsoft.Exchange.WebServices.dll"
        if ($test2 -eq $False) {
            Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] EWS 2.1 not found" -DefaultColor Cyan
            $test3 = Test-Path -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.*\Microsoft.Exchange.WebServices.dll"
            if ($test3 -eq $False) {
                Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] EWS 2.2 not found" -DefaultColor Cyan
            }
            else {
                Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] EWS 2.2 found" -DefaultColor Cyan
            }
        }
        else {
            Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] EWS 2.1 found" -DefaultColor Cyan
        }
    }
    else {
        Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] EWS DLL found in local path" -DefaultColor Cyan
    }

    if ($test -eq $False -and $test2 -eq $False -and $test3 -eq $False) {
        $EwsMessage = @"
You don't seem to have EWS API dll file 'Microsoft.Exchange.WebServices.dll' in the same Directory of this script
please get a copy of the file or download the whole API from:
https://www.microsoft.com/en-us/download/details.aspx?id=42951
        
we will open your browser in 10 seconds automatically directly to this URL
"@
        Write-PSFHostColor -String $EwsMessage -DefaultColor Yellow
        Start-sleep -Seconds 10
        Start-Process "https://www.microsoft.com/en-us/download/details.aspx?id=42951"

        return
    }
    Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] EWS API detected. All good!" -DefaultColor Cyan

    if ($test -eq $True) {
        Add-Type -Path $EWS
        Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] Using EWS DLL in local path" -DefaultColor Cyan
    }
    elseif ($test2 -eq $True) {
        Add-Type -Path "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.*\Microsoft.Exchange.WebServices.dll"
        Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] Using EWS 2.1" -DefaultColor Cyan
    }
    elseif ($test3 -eq $True) {
        Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.*\Microsoft.Exchange.WebServices.dll"
        Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] Using EWS 2.2" -DefaultColor Cyan
    }
}