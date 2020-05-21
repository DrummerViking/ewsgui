@{
	# Script module or binary module file associated with this manifest
	RootModule = 'ewsgui.psm1'
	
	# Version number of this module.
	ModuleVersion = '2.0.9'
	
	# ID used to uniquely identify this module
	GUID = '6a24f2b4-bc88-43fd-9046-19030cf015dc'
	
	# Author of this module
	Author = 'Agustin Gallegos [MSFT]'
	
	# Company or vendor of this module
	CompanyName = 'Agustin Gallegos'
	
	# Copyright statement for this module
	Copyright = 'Copyright (c) 2020 Agustin Gallegos'
	
	# Description of the functionality provided by this module
	Description = 'Exchange Web Services (EWS) tool to perform different operations'
	
	# Minimum version of the Windows PowerShell engine required by this module
	PowerShellVersion = '3.0'
	
	# Modules that must be imported into the global environment prior to importing
	# this module
	RequiredModules = @(
		@{ ModuleName='PSFramework'; ModuleVersion='1.1.59' }
	)
	
	# Assemblies that must be loaded prior to importing this module
	RequiredAssemblies = @('bin\Microsoft.Exchange.WebServices.dll')
	
	# Type files (.ps1xml) to be loaded when importing this module
	# TypesToProcess = @('xml\ewsgui.Types.ps1xml')
	
	# Format files (.ps1xml) to be loaded when importing this module
	# FormatsToProcess = @('xml\ewsgui.Format.ps1xml')
	
	# Functions to export from this module
	FunctionsToExport = 'Start-EWSGui'
	
	# Cmdlets to export from this module
	CmdletsToExport = ''
	
	# Variables to export from this module
	VariablesToExport = ''
	
	# Aliases to export from this module
	AliasesToExport = ''
	
	# List of all modules packaged with this module
	ModuleList = @()
	
	# List of all files packaged with this module
	FileList = @()
	
	# Private data to pass to the module specified in ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
	PrivateData = @{
		
		#Support for PowerShellGet galleries.
		PSData = @{
			
			# Tags applied to this module. These help with module discovery in online galleries.
			# Tags = @()
			
			# A URL to the license for this module.
			LicenseUri = 'https://github.com/agallego-css/ewsgui/blob/master/LICENSE'
			
			# A URL to the main website for this project.
			ProjectUri = 'https://github.com/agallego-css/ewsgui/'
			
			# A URL to an icon representing this module.
			# IconUri = ''
			
			# ReleaseNotes of this module
			# ReleaseNotes = ''
			
		} # End of PSData hashtable
		
	} # End of PrivateData hashtable
}