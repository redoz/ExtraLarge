Set-StrictMode -Version 5

. $PSScriptRoot\New-XLFile.ps1
. $PSScriptRoot\Add-XLSheet.ps1
. $PSScriptRoot\Add-XLTable.ps1
. $PSScriptRoot\Add-XLChart.ps1

Export-ModuleMember -Function New-XLFile
Export-ModuleMember -Function Add-XLSheet
Export-ModuleMember -Function Add-XLTable
Export-ModuleMember -Function Add-XLChart
Export-ModuleMember -Function Add-XLChartSeries

<#
copy C:\dev\ExtraLarge\src -Destination $env:USERPROFILE\Documents\WindowsPowerShell\Modules\ExtraLarge -Container -Recurse -Verbose
Publish-Module -WhatIf -Name ExtraLarge -NuGetApiKey <apikey>
#>