Set-StrictMode -Version 5
Add-Type -Path "$($PSScriptRoot)\EPPlus.dll"

. $PSScriptRoot\New-XLFile.ps1
. $PSScriptRoot\Add-XLSheet.ps1
. $PSScriptRoot\Add-XLTable.ps1
. $PSScriptRoot\Add-XLChart.ps1

Export-ModuleMember -Function New-XLFile
Export-ModuleMember -Function Add-XLSheet
Export-ModuleMember -Function Add-XLTable
Export-ModuleMember -Function Add-XLChart
Export-ModuleMember -Function Add-XLChartSeries

              