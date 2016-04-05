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
How to generate module manifest:
New-ModuleManifest -ModuleVersion 0.0.1 -Path .\src\ExtraLarge.psd1 -RootModule ExtraLarge.psm1 -FileList (Get-ChildItem -Path src).Name -RequiredAssemblies "EPPlus.dll" -Author "Patrik Husfloen" -CompanyName "ExtraLarge" -DefaultCommandPrefix "XL" -ProcessorArchitecture MSIL -DotNetFrameworkVersion 3.5 -Tags Excel,EPPlus,Charts -ProjectUri http://github.com/redoz/ExtraLarge -Description "Create Excel files using PowerShell"
copy C:\dev\vendor\ExtraLarge\src\* -Destination $env:USERPROFILE\Documents\WindowsPowerShell\Modules\ExtraLarge -Container -Recurse
Publish-Module -WhatIf -Name ExtraLarge -NuGetApiKey <apikey>
#>

              