Set-StrictMode -Version 5

$module = Get-Module -Name ExtraLarge
$modulePath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath ..\src))

if (-not $module -or $module.ModuleBase -ne $modulePath) {
    Import-Module -Name (Join-Path -Path $modulePath -ChildPath ExtraLarge.psd1) -Force -Verbose
}

function Get-TestPath {
    param(
        $FileName = "Test.xlsx"
    )
    
    [string]$path = Join-Path -Path ((Get-PSDrive TestDrive).Root) -ChildPath $FileName;
    if (Test-Path -Path $path) {
        Remove-Item -Path $path -Force
    }
    
    return $path
}   