. .\test\Setup.ps1

function Test-Range {
    param ($Sheet)
    $worksheet = $Sheet.Worksheet
    $worksheet.MergedCells | Where-Object -FilterScript {$_ -ne $null} | Should Be $null
}

Describe "Split-XLRange/Sheet" {
    Context "With -PassThru" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
                    
        It "Should return [XLSheet]" {
            Join-XLRange -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4
            Split-XLRange -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4 -PassThru | %{$_ -is [XLSheet]} | Should Be $true
            Test-Range -Sheet $sheet
        }
    }
    Context "Without -PassThru" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
                    
        It "Should return [XLRange]" {
            Join-XLRange -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4 
            Split-XLRange -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4 |  %{$_ -is [XLRange]} | Should Be $true
            Test-Range -Sheet $sheet
        }
    }
}

Describe "Splt-XLRange/Range" {
        Context "With -PassThru" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
                    
        It "Should return [XLRange]" {
            $range = Select-XLRange -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4
            Join-XLRange -Range $range -PassThru
            Split-XLRange -Range $range -PassThru | %{$_ -is [XLRange]} | Should Be $true
            Test-Range -Sheet $sheet       
        }
    }
    Context "Without -PassThru" {
        [string]$path = Get-TestPath
        $sheet = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X"
                    
        It "Should return [XLRange]" {
            $range = Select-XLRange -Sheet $sheet -FromColumn 2 -ToColumn 4 -FromRow 2 -ToRow 4
            Join-XLRange -Range $range 
            Split-XLRange -Range $range | %{$_ -is [XLRange]} | Should Be $true
            Test-Range -Sheet $sheet
        }
    }
}
