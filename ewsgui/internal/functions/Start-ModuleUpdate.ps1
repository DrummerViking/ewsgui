Function Start-ModuleUpdate {
	<#
	.SYNOPSIS
	Function to start checking for updates on this module.
	
	.DESCRIPTION
	Function to start checking for updates on this module.
	
	.PARAMETER ModuleRoot
	Modules root path.

	.PARAMETER Confirm
    If this switch is enabled, you will be prompted for confirmation before executing any operations that change state.

    .PARAMETER WhatIf
    If this switch is enabled, no actions are performed but informational messages will be displayed that explain what would happen if the command were to run.
	
	.EXAMPLE
	PS C:\> Start-ModuleUpdate -ModuleRoot "C:\Temp"
	Runs the function to start checking for update for current module in "C:\Temp"
	#>
	[CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'Low')]
	Param (
		[String]$ModuleRoot
	)

	$ScriptBlock = {
		Param (
			[String]$ModuleRoot
		)
		$moduleManifest = (Import-PowerShellDataFile -Path "$((Get-ChildItem -Path $ModuleRoot -Filter *.psd1).Fullname)")
		$moduleFileName = $moduleManifest.RootModule
		$moduleName = $ModuleFileName.Substring(0, $ModuleFileName.IndexOf("."))
		$script:ModuleVersion = $moduleManifest.ModuleVersion -as [version]

		$GalleryModule = Find-Module -Name $ModuleName -Repository PSGallery
		if ( $script:ModuleVersion -lt $GalleryModule.version ) {
			$bt = New-BTButton -Content "Get Update" -Arguments "$($moduleManifest.PrivateData.PSData.ProjectUri)#installation"
			New-BurntToastNotification -Text "$ModuleName Update found", "There is a new version $($GalleryModule.version) of this module available." -Button $bt
		}
	}

	# Create Runspace, set maximum threads
	$pool = [RunspaceFactory]::CreateRunspacePool(1, 1)
	$pool.ApartmentState = "MTA"
	$pool.Open()

	$runspace = [PowerShell]::Create()
	$runspace.Runspace.Name = "$ModuleName.Update"
	$null = $runspace.AddScript( $ScriptBlock )
	$null = $runspace.AddArgument( $ModuleRoot )
	$runspace.RunspacePool = $pool

	[PSCustomObject]@{
		Pipe   = $runspace
		Status = $runspace.BeginInvoke()
		Pool   = $pool
	}
}