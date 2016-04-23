Function Save-XLFile {
[CmdletBinding()]
[OutputType([XLFile], ParameterSetName = "File")]
[OutputType([XLSheet], ParameterSetName = "Sheet")]
param(
    [Parameter(ParameterSetName = "File", Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
    [XLFile]$File,
    [Parameter(ParameterSetName = "Sheet", Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
    [XLSheet]$Sheet,
    
    [switch]$PassThru = $false
)
begin{}
process{
    $package = $null
    if ($PSCmdlet.ParameterSetName -eq "File") {
        $package = $File.Package;
    } else {
        $package = $Sheet.Owner
    }
    
    if ($package -eq $null)
    {
        throw "Unable to save, no ExcelPackage found."
    }
    
    $package.Save();
    
    if ($PassThru.IsPresent) {
        if ($PSCmdlet.ParameterSetName -eq "File") {
            Write-Output -InputObject $File
        } else {
            Write-Output -InputObject $Sheet
        }
    } 
}
end {
}
}
