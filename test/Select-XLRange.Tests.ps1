. .\test\Setup.ps1

Describe "Select-XLRange" {
    Context "With table data" {
        [string]$path = Get-TestPath
        $data = ConvertFrom-Csv -InputObject "A;B;C`n1;2;3;`n101;102;103" -Delimiter ';'
        $file = New-XLFile -Path $path -PassThru -With { $_ |
                    Add-XLSheet -Name "X" |
                    Add-XLTable -Name "T" -Row 2 -Column 2 -Data $data }
                    
        It "Should return table data with headers" {
            $range = Select-XLRange -Sheet (Get-XLSheet $file) -HasHeaders -FromRow 2 -FromColumn 2 -ToRow 4 -ToColumn 4
            $range -is [XLRange] | Should Be $true
            $rows = ($range | Select-Object -First 4)
            $rows.Count | Should Be 2
            $rows[0].A | Should Be 1
            $rows[0].B | Should Be 2
            $rows[0].C | Should Be 3
            $rows[1].A | Should Be 101
            $rows[1].B | Should Be 102
            $rows[1].C | Should Be 103
        }
        
        It "Should return table data with replaced headers" {
            $range = Select-XLRange -Sheet (Get-XLSheet $file) -HasHeaders -FromRow 2 -FromColumn 2 -ToRow 4 -ToColumn 4 -Headers X,Y,Z
            $range -is [XLRange] | Should Be $true
            $rows = ($range | Select-Object -First 4)
            $rows.Count | Should Be 2
            $rows[0].X | Should Be 1
            $rows[0].Y | Should Be 2
            $rows[0].Z | Should Be 3
            $rows[1].X | Should Be 101
            $rows[1].Y | Should Be 102
            $rows[1].Z | Should Be 103
        }
        It "Should return table data using excel column letters" {
            $range = Select-XLRange -Sheet (Get-XLSheet $file) -FromRow 3 -FromColumn 3 -ToRow 4 -ToColumn 4
            $range -is [XLRange] | Should Be $true
            $rows = ($range | Select-Object -First 4)
            $rows.Count | Should Be 2
            $rows[0].C | Should Be 2
            $rows[0].D | Should Be 3
            $rows[1].C | Should Be 102
            $rows[1].D | Should Be 103
        }        
        It "Should throw if -Headers count doesn't match the column count" {
            { Select-XLRange -Sheet (Get-XLSheet $file) -HasHeaders -FromRow 2 -FromColumn 2 -ToRow 4 -ToColumn 4 -Headers X,Y,Z,T } | Should Throw
        }                 
    }
}
