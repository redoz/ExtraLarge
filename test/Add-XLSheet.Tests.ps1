. .\test\Setup.ps1

Describe "Add-XLSheet" {
    Context "Worksheet already exists" {
        [string]$path = Get-TestPath
        $xl = New-XLFile -Path $path -PassThru | Add-XLSheet -Name "X" -PassThru
                    
        It "should throw without -Force" {
            { $xl | Add-XLSheet -Name "X" } | Should Throw
        }
        
        It "should overwrite with -Force" {
            $old = $xl.Package.Workbook.Worksheets["X"]
            $new = $xl | Add-XLSheet -Name "X" -Force -PassThru
            [object]::ReferenceEquals($old, $new) | Should Be $false
        }
    }
    Context "Worksheet doesn't exist" {
        [string]$path = Get-TestPath
        $file = New-XLFile -Path $path -PassThru        
            
        It "Should return an [XLFile] with -PassThru" {    
            Add-XLSheet -File $file -Name "X" -PassThru:$true | %{$_ -is [XLFile]} | Should Be $true
            $path | Should Exist
        }
    
        It "should return an [XLSheet] without -PassThru" {         
            Add-XLSheet -File $file -Name "Y" -PassThru:$false | %{$_ -is [XLSheet]} | Should Be $true
            $path | Should Exist
        }
    }
    
    Context "Passing file as Path" {
        It "Should throw if file does not exist" {
             { Add-XLSheet -Path (Get-TestPath -FileName 'NoFile.xslx') -Name "Y" } | Should Throw
        }
        
        It "Should work if file exists" {
            [string]$path = Get-TestPath
            $xl = New-XLFile -Path $path
                  
            Add-XLSheet -Path $path -Name "X" | %{$_ -is [XLSheet]} | Should Be $true
            $path | Should Exist
            
#            [OfficeOpenXml.ExcelPackage]::new($path)
        }
    }
}