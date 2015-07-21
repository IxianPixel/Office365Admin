@{
	## Module Info
	ModuleVersion      = '0.0.0.1'
	Description        = 'Office365Admin Module'
	GUID               = '77678c53-a074-4995-a829-fc82b75a09a3'

	## Module Components
	ScriptsToProcess   = @()
	ModuleToProcess    = @("Office365Admin.psm1")
	TypesToProcess     = @()
	FormatsToProcess   = @()
	ModuleList         = @("Office365Admin.psm1")
	FileList           = @()

	## Public Interface
	CmdletsToExport    = ''
	FunctionsToExport  = '*'
	VariablesToExport  = '*'
	AliasesToExport    = '*'

	## Requirements
	PowerShellVersion      = '3.0'
	PowerShellHostName     = ''
	PowerShellHostVersion  = ''
	RequiredModules        = @()
	RequiredAssemblies     = @()
	ProcessorArchitecture  = 'None'
	DotNetFrameworkVersion = '2.0'
	CLRVersion             = '2.0'

	## Author
	Author             = 'Dylan Addison'
	CompanyName        = 'Wavehill'
	Copyright          = ''

	## Private Data
	PrivateData        = ''
}
