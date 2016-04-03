Import-Module -Name "${PSScriptRoot}\..\ExtraLarge\ExtraLarge.psm1" -Force
Describe "New-XLFile" {
    Context "There are no worksheets added" {
        It "creates a default worksheet" {
            [string]$path = "c:\temp\test.xslx";
            if (Test-Path -Path $path) {
                Remove-Item -Path $path -Force
            }
            New-XLFile -Path $path | Should BeNullOrEmpty
            $path | Should Exist
        }
    }
    Context "PassThru is set" {
        It "should return an [OfficeOpenXml.ExcelPackage]" {
            [string]$path = "c:\temp\test.xslx";
            if (Test-Path -Path $path) {
                Remove-Item -Path $path -Force
            }
            New-XLFile -Path $path -PassThru | %{$_ -is [OfficeOpenXml.ExcelPackage]} | Should Be $true
            $path | Should Exist
        }
    }
    
    Context "PassThru is not set" {
        It "should return nothing" {
            [string]$path = "c:\temp\test.xslx";
            if (Test-Path -Path $path) {
                Remove-Item -Path $path -Force
            }
            New-XLFile -Path $path -PassThru:$false | Should BeNullOrEmpty
            $path | Should Exist
        }
    }
    
    Context "File exists" {
        It "should throw without -Force" {
            [string]$path = "c:\temp\test.xslx";
            Set-Content -Path $path -Value "" -Force
            $path | Should Exist
            { New-XLFile -Path $path } | Should Throw
            $path | Should Exist
        }
        It "should overwrite with -Force" {
            [string]$path = "c:\temp\test.xslx";
            Set-Content -Path $path -Value @() -Encoding Byte -Force
            $path | Should Exist
            (Get-Item -Path $path).Length -eq 0 | Should Be $true
            New-XLFile -Path $path -Force | Should BeNullOrEmpty
            (Get-Item -Path $path).Length -gt 0 | Should Be $true
        }
    }
}

Describe "Add-XLSheet" {
    Context "Worksheet exists" {
        It "should throw without -Force" {
            [string]$path = "c:\temp\test.xslx";
            if (Test-Path -Path $path) {
                Remove-Item -Path $path -Force
            }
            $xl = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X" -PassThru
            { $xl | Add-XLSheet -Name "X" } | Should Throw
        }
        It "should overwrite with -Force" {
            [string]$path = "c:\temp\test.xslx";
            if (Test-Path -Path $path) {
                Remove-Item -Path $path -Force
            }
            $xl = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X" -PassThru
            $old = $xl.Workbook.Worksheets["X"]
            $new = $xl | Add-XLSheet -Name "X" -Force -PassThru
            [object]::ReferenceEquals($old, $new) | Should Be $false
        }
    }
    Context "PassThru is set" {
        It "should return an [OfficeOpenXml.ExcelPackage]" {
            [string]$path = "c:\temp\test.xslx";
            if (Test-Path -Path $path) {
                Remove-Item -Path $path -Force
            }
            New-XLFile -Path $path -PassThru | 
                Add-XLSheet -Name "X" -PassThru:$true | %{$_ -is [OfficeOpenXml.ExcelPackage]} | Should Be $true
            $path | Should Exist
        }
    }
    
    Context "PassThru is not set" {
        It "should return an [OfficeOpenXml.ExcelWorksheet]" {
            [string]$path = "c:\temp\test.xslx";
            if (Test-Path -Path $path) {
                Remove-Item -Path $path -Force
            }
            New-XLFile -Path $path -PassThru | 
                Add-XLSheet -Name "X" -PassThru:$false | %{$_ -is [OfficeOpenXml.ExcelWorksheet]} | Should Be $true
            $path | Should Exist
        }
    }
}