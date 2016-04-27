. .\test\Setup.ps1

function Test-Range {
    param ($Sheet)
    $worksheet = $Sheet.Worksheet
    $worksheet.Cells[2,2].Value | Should Be 99
    $worksheet.Cells[2,3].Value | Should Be 99
    $worksheet.Cells[2,4].Value | Should Be 99
    $worksheet.Cells[3,2].Value | Should Be 99
    $worksheet.Cells[3,3].Value | Should Be 99
    $worksheet.Cells[3,4].Value | Should Be 99
    $worksheet.Cells[4,2].Value | Should Be 99
    $worksheet.Cells[4,3].Value | Should Be 99
    $worksheet.Cells[4,4].Value | Should Be 99
}

Describe "Set-XLValue/Sheet" {
    Context "With -PassThru" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
                    
        It "Should return [XLSheet]" {
            $ret = Set-XLValue -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4 -Value 99 -PassThru
            $ret -is [XLSheet] | Should Be $true
            Test-Range -Sheet $sheet
        }
    }
    Context "Without -PassThru" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
                    
        It "Should return [XLRange]" {
            $ret = Set-XLValue -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4 -Value 99 
            $ret -is [XLRange] | Should Be $true
            Test-Range -Sheet $sheet            
        }
    }
}

Describe "Set-XLValue/Range" {
        Context "With -PassThru" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
                    
        It "Should return [XLRange]" {
            $range = Select-XLRange -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4
            $ret = Set-XLValue -Range $range -Value 99 -PassThru
            $ret -is [XLRange] | Should Be $true
            Test-Range -Sheet $sheet       
        }
    }
    Context "Without -PassThru" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
                    
        It "Should return nothing" {
            $range = Select-XLRange -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4
            $ret = Set-XLValue -Range $range -Value 99        
            $ret -is [XLRange] | Should Be $true
            Test-Range -Sheet $sheet
        }
    }  
}


Describe "Set-XLValue/Named" {
    Context "With -PassThru" {
        [string]$path = Get-TestPath
        Copy-Item -Path (Join-Path -Path $PSScriptRoot -ChildPath "data\WithNamedRange.xlsx") -Destination $path
        $sheet = Get-XLFile -Path $path | Get-XLSheet -Index 1      
        It "Should return [XLSheet]" {
            $ret = Set-XLValue -Sheet $sheet -Name "Name" -Value 99 -PassThru 
            $ret -is [XLSheet] | Should Be $true
            Test-Range -Sheet $sheet
        }
    }
    Context "Without -PassThru" {
        [string]$path = Get-TestPath
        Copy-Item -Path (Join-Path -Path $PSScriptRoot -ChildPath "data\WithNamedRange.xlsx") -Destination $path
        $sheet = Get-XLFile -Path $path | Get-XLSheet -Index 1
                    
        It "Should return [XLRange]" {
            $ret = Set-XLValue -Sheet $sheet -Name "Name" -Value 99 
            $ret -is [XLRange] | Should Be $true
            Test-Range -Sheet $sheet            
        }
    } 
}
