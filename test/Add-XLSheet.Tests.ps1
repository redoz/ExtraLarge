BeforeAll { 
    . $PSScriptRoot\Setup.ps1
}

Describe "Add-XLSheet/File" {
    Context "Worksheet already exists" {
        BeforeAll {
            [string]$path = Get-TestPath
            $xl = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X" -PassThru
        }
                    
        It "should throw without -Force" {
            { $xl | Add-XLSheet -Name "X" } | Should -Throw
        }
        
        It "should overwrite with -Force" {
            $old = $xl.Package.Workbook.Worksheets["X"]
            $new = $xl | Add-XLSheet -Name "X" -Force -PassThru
            [object]::ReferenceEquals($old, $new) | Should -Be $false
        }
    }
    Context "Worksheet doesn't exist" {
        BeforeAll {
            [string]$path = Get-TestPath
            $file = New-XLFile -Path $path -PassThru        
        }
            
        It "Should return an [XLFile] with -PassThru" {    
            Add-XLSheet -File $file -Name "X" -PassThru:$true | %{$_ -is [XLFile]} | Should -Be $true
            $path | Should -Exist
        }
    
        It "should return an [XLSheet] without -PassThru" {         
            Add-XLSheet -File $file -Name "Y" -PassThru:$false | %{$_ -is [XLSheet]} | Should -Be $true
            $path | Should -Exist
        }
    }
}
Describe "Add-XLSheet/Path" {    
    Context "File doesn't exist" {
        It "Should throw if file does not exist" {
             { Add-XLSheet -Path (Get-TestPath -FileName 'NoFile.xslx') -Name "Y" } | Should -Throw
        }
    }
    Context "File exists" {
        BeforeAll {
            [string]$path = Get-TestPath
            $xl = New-XLFile -Path $path
        }
                    
        It "Should return an [XLSheet] without -PassThru" {
            Add-XLSheet -Path $path -Name "X" -Save | %{$_ -is [XLSheet]} | Should -Be $true
            $path | Should -Exist
            
            $package = [OfficeOpenXml.ExcelPackage]::new($path)
            $package.Workbook.Worksheets["X"] | Should -Not -Be $null
        }
        
        It "Should return an [XLFile] with -PassThru" {
            Add-XLSheet -Path $path -Name "Y" -PassThru -Save | %{$_ -is [XLFile]} | Should -Be $true
            $path | Should -Exist
            
            $package = [OfficeOpenXml.ExcelPackage]::new($path)
            $package.Workbook.Worksheets["Y"] | Should -Not -Be $null
        }
    }
}
