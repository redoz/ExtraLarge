Set-StrictMode -Version 5

$typeDefinitionPath = Join-Path -Path $PSScriptRoot -ChildPath 'Types.cs'
[string]$typeDefinitions = Get-Content -Path $typeDefinitionPath -Raw -Encoding UTF8

$epplusPath = Join-Path -Path $PSScriptRoot -ChildPath 'EPPlus.dll'

Add-Type -TypeDefinition $typeDefinitions -ReferencedAssemblies $epplusPath

. $PSScriptRoot\New-XLFile.ps1
. $PSScriptRoot\Add-XLSheet.ps1
. $PSScriptRoot\Add-XLTable.ps1
. $PSScriptRoot\Add-XLChart.ps1

. $PSScriptRoot\Get-XLSheet.ps1
. $PSScriptRoot\Get-XLFile.ps1

Export-ModuleMember -Function New-XLFile
Export-ModuleMember -Function Add-XLSheet
Export-ModuleMember -Function Add-XLTable
Export-ModuleMember -Function Add-XLChart
Export-ModuleMember -Function Add-XLChartSeries
Export-ModuleMember -Function Get-XLSheet
Export-ModuleMember -Function Get-XLFile

<#
copy C:\dev\ExtraLarge\src -Destination $env:USERPROFILE\Documents\WindowsPowerShell\Modules\ExtraLarge -Container -Recurse -Verbose
Publish-Module -WhatIf -Name ExtraLarge -NuGetApiKey <apikey>
#>