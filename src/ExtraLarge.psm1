Set-StrictMode -Version 5


$typeDefinitionPath = Join-Path -Path $PSScriptRoot -ChildPath 'Types.cs'
[string]$typeDefinitions = Get-Content -Path $typeDefinitionPath -Raw -Encoding UTF8

$epplusPath = Join-Path -Path $PSScriptRoot -ChildPath 'EPPlus.dll'
try {
    Add-Type -TypeDefinition $typeDefinitions -ReferencedAssemblies $epplusPath
} catch  {
    Write-Warning -Message ('Failed to load ExtraLarge types: ' + $_)
}

. $PSScriptRoot\New-XLFile.ps1
. $PSScriptRoot\Add-XLSheet.ps1
. $PSScriptRoot\Add-XLTable.ps1
. $PSScriptRoot\Add-XLChart.ps1

. $PSScriptRoot\Get-XLSheet.ps1
. $PSScriptRoot\Get-XLFile.ps1
. $PSScriptRoot\Save-XLFile.ps1

. $PSScriptRoot\Select-XLRange.ps1
. $PSScriptRoot\Join-XLRange.ps1
. $PSScriptRoot\Split-XLRange.ps1
. $PSScriptRoot\Clear-XLRange.ps1

Export-ModuleMember -Function New-XLFile
Export-ModuleMember -Function Add-XLSheet
Export-ModuleMember -Function Add-XLTable
Export-ModuleMember -Function Add-XLChart
Export-ModuleMember -Function Add-XLChartSeries

Export-ModuleMember -Function Get-XLSheet
Export-ModuleMember -Function Get-XLFile
Export-ModuleMember -Function Save-XLFile

Export-ModuleMember -Function Select-XLRange
Export-ModuleMember -Function Join-XLRange
Export-ModuleMember -Function Split-XLRange
Export-ModuleMember -Function Clear-XLRange
<#
copy C:\dev\ExtraLarge\src -Destination $env:USERPROFILE\Documents\WindowsPowerShell\Modules\ExtraLarge -Container -Recurse -Verbose
Publish-Module -WhatIf -Name ExtraLarge -NuGetApiKey <apikey>
#>
