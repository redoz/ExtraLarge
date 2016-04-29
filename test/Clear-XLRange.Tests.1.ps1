. .\test\Setup.ps1

function Test-Range {
    param ($Sheet)
    $worksheet = $Sheet.Worksheet
    $worksheet.MergedCells.Count | Should Be 1
    $address = [OfficeOpenXml.ExcelAddress]$worksheet.MergedCells[0]
    $address.Start.Row | Should Be 2
    $address.Start.Column | Should Be 2
    $address.End.Row | Should Be 4
    $address.End.Column | Should Be 4          
}

Describe "Clear-XLRange" {
    Context "With -PassThru" {
        $path = Join-Path $TestDrive 'test.xlsx'
        Copy-Item -Path .\test\data\WithNamedRange.xlsx -Destination $path
        $xlFile = Get-XLFile -Path $path 
        
        $res = Clear-XLRange -File $xlFile -Address -PassThru "Sheet2!B2:B4"
                    
        It "Should return [XLFile]" {
            $res -is [XLFile] | Should Be $true
        }
        
        It "Should have cleared data" {
            $xlSheet = $xlFile | Get-XLSheet -Index 2
            $xlSheet.Worksheet.Cells["B2"].Value | Should Be ''
            $xlSheet.Worksheet.Cells["B3"].Value | Should Be ''
            $xlSheet.Worksheet.Cells["B4"].Value | Should Be ''
        }
    }
    Context "Without -PassThru" {
        $path = Join-Path $TestDrive 'test.xlsx'
        Copy-Item -Path .\test\data\WithNamedRange.xlsx -Destination $path
        $xlFile = Get-XLFile -Path $path 
        
        $res = Clear-XLRange -File $xlFile -Address "Sheet2!B2:B4"
                    
        It "Should return [XLRange]" {
            $res -is [XLRange] | Should Be $true
        }
        
        It "Should have cleared data" {
            $xlSheet = $xlFile | Get-XLSheet -Index 2
            $xlSheet.Worksheet.Cells["B2"].Value | Should Be ''
            $xlSheet.Worksheet.Cells["B3"].Value | Should Be ''
            $xlSheet.Worksheet.Cells["B4"].Value | Should Be ''
        }
    }
}

