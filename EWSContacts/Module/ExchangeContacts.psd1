@{
	
	# Script module or binary module file associated with this manifest
	ModuleToProcess = 'ExchangeContacts.psm1'
	
	# Version number of this module.
	ModuleVersion = '1.6.0.0'
	
	# ID used to uniquely identify this module
	GUID = '6901af86-a39e-4a4c-afa6-b8ead127c674'
	
	# Author of this module
	Author = 'Glen Scales'
	
	# Company or vendor of this module
	CompanyName = ''
	
	# Copyright statement for this module
	Copyright = '(c) 2017. All rights reserved.'
	
	# Description of the functionality provided by this module
	Description = 'A module to manage contacts in Exchange using EWS'
	
	# Minimum version of the Windows PowerShell engine required by this module
	PowerShellVersion = '2.0'
	
	# Name of the Windows PowerShell host required by this module
	PowerShellHostName = ''
	
	# Minimum version of the Windows PowerShell host required by this module
	PowerShellHostVersion = ''
	
	# Minimum version of the .NET Framework required by this module
	DotNetFrameworkVersion = '2.0'
	
	# Minimum version of the common language runtime (CLR) required by this module
	CLRVersion = '2.0.50727'
	
	# Processor architecture (None, X86, Amd64, IA64) required by this module
	ProcessorArchitecture = 'None'
	
	# Modules that must be imported into the global environment prior to importing
	# this module
	RequiredModules = @()
	
	# Assemblies that must be loaded prior to importing this module
	RequiredAssemblies = @()
	
	# Script files (.ps1) that are run in the caller's environment prior to
	# importing this module
	ScriptsToProcess = @()
	
	# Type files (.ps1xml) to be loaded when importing this module
	TypesToProcess = @()
	
	# Format files (.ps1xml) to be loaded when importing this module
	FormatsToProcess = @()
	
	# Modules to import as nested modules of the module specified in
	# ModuleToProcess
	NestedModules = @()
	
	# Functions to export from this module
	FunctionsToExport	    = @(
		'Connect-EXCExchange',
		'Copy-EXCContactGalToMailbox',
		'Export-EXCContact',
		'Export-EXCContacts',
		'Export-EXCRootContacts',
		'Export-EXCContactFolder',
		'Export-EXCGALContact',
		'Export-EXCDumpsterContacts',
		'Get-EXCContact',
		'Get-EXCContactFolder',
		'Get-EXCContactGroup',
		'Get-EXCContacts',
		'Get-EXCAllContactFolders',
		'Get-EXCContactGroups',
		'Get-EXCDumpsterContacts',
		'New-EXCContact',
		'New-EXCContactGroup',
		'Get-EXCPrivacyReport',
		'Remove-EXCContact',
		'Search-EXCContactsForCCNumbers',
		'Search-EXCAllContactFolders',
		'Search-EXCAllContactGroups',
		'Invoke-ContactToCSVEntry',
		'Set-EXCContact'
	)
	
	# Cmdlets to export from this module
	CmdletsToExport = '' 
	
	# Variables to export from this module
	VariablesToExport = ''
	
	# Aliases to export from this module
	AliasesToExport = '*' #For performance, list alias explicity
	
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
			# LicenseUri = 'https://github.com/gscales/Powershell-Scripts/tree/master/LICENSE.txt'
			
			# A URL to the main website for this project.
			# ProjectUri = 'https://github.com/gscales/Powershell-Scripts/tree/master/EWSContacts'
			
			# A URL to an icon representing this module.
			# IconUri = ''
			
			# ReleaseNotes of this module
			# ReleaseNotes = ''
			
		} # End of PSData hashtable
		
	} # End of PrivateData hashtable
}







