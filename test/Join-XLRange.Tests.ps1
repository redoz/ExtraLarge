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

Describe "Join-XLRange/Sheet" {
    Context "With -PassThru" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
                    
        It "Should return [XLSheet]" {
            Join-XLRange -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4 -PassThru | %{$_ -is [XLSheet]} | Should Be $true
            Test-Range -Sheet $sheet
        }
    }
    Context "Without -PassThru" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
                    
        It "Should return [XLRange]" {
            Join-XLRange -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4 | %{$_ -is [XLRange]} | Should Be $true
            Test-Range -Sheet $sheet            
        }
    }
    
    Context "Overlapping with existing table" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
        Add-XLTable -Name "TBL1" -Sheet $sheet -Row 2 -Column 2 -Data (ConvertFrom-Csv -Delimiter ';' -InputObject "A;B`n1;2")
        It "Should throw" {
            { Join-XLRange -Sheet $sheet -FromRow 2 -FromColumn 2 -ToRow 4 -ToColumn 4 } | Should Throw
        }
    }
}

Describe "Join-XLRange/Range" {
        Context "With -PassThru" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
                    
        It "Should return [XLRange]" {
            $range = Select-XLRange -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4
            Join-XLRange -Range $range -PassThru | %{$_ -is [XLRange]} | Should Be $true
            Test-Range -Sheet $sheet       
        }
    }
    Context "Without -PassThru" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
                    
        It "Should return [XLRange]]" {
            $range = Select-XLRange -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4
            Join-XLRange -Range $range | %{$_ -is [XLRange]} | Should Be $true
            Test-Range -Sheet $sheet
        }
    }
    
    Context "Overlapping with existing table" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
        Add-XLTable -Name "TBL1" -Sheet $sheet -Row 2 -Column 2 -Data (ConvertFrom-Csv -Delimiter ';' -InputObject "A;B`n1;2")
        It "Should throw" {
            $range = Select-XLRange -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4
            { Join-XLRange -Range $range} | Should Throw
        }
    }    
}
