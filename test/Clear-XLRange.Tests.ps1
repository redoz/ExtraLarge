BeforeAll { 
    . $PSScriptRoot\Setup.ps1
}

Describe "Clear-XLRange" {
    Context "With -PassThru" {
        BeforeAll {
            $path = Join-Path $TestDrive 'test.xlsx'
            Copy-Item -Path $PSScriptRoot\data\WithNamedRange.xlsx -Destination $path
            $xlFile = Get-XLFile -Path $path 
            
            $res = Clear-XLRange -File $xlFile -Address  "Sheet2!B2:B4" -PassThru
        }
                    
        It "Should return [XLFile]" {
            $res -is [XLFile] | Should -Be $true
        }
        
        It "Should have cleared data" {
            $xlSheet = $xlFile | Get-XLSheet -Index 1
            $xlSheet.Worksheet.Cells["B2"].Value | Should -Be $null
            $xlSheet.Worksheet.Cells["B3"].Value | Should -Be $null
            $xlSheet.Worksheet.Cells["B4"].Value | Should -Be $null
        }
    }
    Context "Without -PassThru" {
        BeforeAll {
            $path = Join-Path $TestDrive 'test.xlsx'
            Copy-Item -Path $PSScriptRoot\data\WithNamedRange.xlsx -Destination $path
            $xlFile = Get-XLFile -Path $path 
            
            $res = Clear-XLRange -File $xlFile -Address "Sheet2!B2:B4"
        }
                    
        It "Should return [XLRange]" {
            $res -is [XLRange] | Should -Be $true
        }
        
        It "Should have cleared data" {
            $xlSheet = $xlFile | Get-XLSheet -Index 1
            $xlSheet.Worksheet.Cells["B2"].Value | Should -Be $null
            $xlSheet.Worksheet.Cells["B3"].Value | Should -Be $null
            $xlSheet.Worksheet.Cells["B4"].Value | Should -Be $null
        }
    }
}

