Function Save-XLFile {
[CmdletBinding()]
[OutputType([XLFile], ParameterSetName = "File")]
[OutputType([XLSheet], ParameterSetName = "Sheet")]
param(
    [Parameter(ParameterSetName = "File", Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
    [XLFile]$File,
    [Parameter(ParameterSetName = "Sheet", Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
    [XLSheet]$Sheet,
    [System.IO.FileInfo]$Path,
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
    
    if ($null -eq $package)
    {
        throw "Unable to save, no ExcelPackage found."
    }
    
    if ($null -ne $Path) {
        $package.SaveAs($Path);
    } else {
        $package.Save();
    }
    
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
