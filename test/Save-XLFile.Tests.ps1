. .\test\Setup.ps1

Describe "Save-XLFile/File" {
    Context "File already exists" {
        [string]$path = Get-TestPath
        $xl = New-XLFile -Path $path -PassThru
        
        $lastWritTimeBefore = (Get-Item -Path $path).LastWriteTime
        It "Should be updated" {
            $path | Should Exist
            Start-Sleep -Milliseconds 500
            Save-XLFile -File $xl
            $fileAfter = Get-Item -Path $path
            $path | Should Exist
            $lastWritTimeAfter = (Get-Item -Path $path).LastWriteTime
            $lastWritTimeBefore -ne $lastWritTimeAfter | Should Be $true
        }
    }
    Context "File doesn't exist" {
        [string]$path = Get-TestPath
        $file = New-XLFile -Path $path -PassThru -NoDefaultSheet
            
        It "Should return be created" {
            # add a sheet so it can be saved 
            $null = Add-XLSheet -File $file -Name "X"
            $path | Should Not Exist
            Save-XLFile -File $file
            $path | Should Exist
        }
    }
    Context "With -PassThru" {
        [string]$path = Get-TestPath
        $file = New-XLFile -Path $path -PassThru
        It "Should return [XLFile]" {
            Save-XLFile -File $file -PassThru | %{$_ -is [XLFile]} | Should Be $true
        }
    }
}

Describe "Save-XLFile/Sheet" {
    Context "File already exists" {
        [string]$path = Get-TestPath
        $xl = New-XLFile -Path $path -PassThru
        
        $lastWritTimeBefore = (Get-Item -Path $path).LastWriteTime
        It "Should be updated" {
            $path | Should Exist
            Start-Sleep -Milliseconds 500
            Save-XLFile -Sheet (Get-XLSheet -File $xl -Index 1)
            $fileAfter = Get-Item -Path $path
            $path | Should Exist
            $lastWritTimeAfter = (Get-Item -Path $path).LastWriteTime
            $lastWritTimeBefore | Should Not Be $lastWritTimeAfter
        }
    }
    Context "File doesn't exist" {
        [string]$path = Get-TestPath
        $file = New-XLFile -Path $path -PassThru -NoDefaultSheet
            
        It "Should return be created" {
            # add a sheet so it can be saved 
            $null = Add-XLSheet -File $file -Name "X"
            $path | Should Not Exist
            Save-XLFile -Sheet (Get-XLSheet -File $file -Index 1)
            $path | Should Exist
        }
    }
    Context "Sheet has no owner" {
        [string]$path = Get-TestPath
        $file = New-XLFile -Path $path -PassThru
            
        It "Should throw" {
            [XLSheet]$sheet = (Get-XLSheet -File $file -Index 1).Worksheet
            { Save-XLFile -Sheet $sheet } | Should Throw
        }
    }
    Context "With -PassThru" {
        [string]$path = Get-TestPath
        $file = New-XLFile -Path $path -PassThru
        It "Should return [XLSheet]" {
            Save-XLFile -Sheet (Get-XLSheet -File $file -Index 1) -PassThru | %{$_ -is [XLSheet]} | Should Be $true
        }
    }
}
