param(
    [Parameter(Mandatory=$true)]
    $RepoToken,
    [Parameter(Mandatory=$true)]
    $ServiceName,
    [Parameter(Mandatory=$true)]
    $ServiceJobId,
    [Parameter(Mandatory=$true)]
    $CodeCoverage)

# prepare coverage report
$fileReports = [System.Collections.Generic.List[PSObject]]::new()
foreach ($file in $CodeCoverage.AnalyzedFiles) {
    Write-Verbose -Message "Processing coverage for: $file" -Verbose
    $fileName = [System.IO.Path]::GetFileName($file)
    $digest = Get-FileHash -LiteralPath $file -Algorithm MD5
    
    $fileHits = $CodeCoverage.HitCommands.Where({$_.File -eq $file})
    $fileMisses = $CodeCoverage.MissedCommands.Where({$_.File -eq $file})

    $lineCount = (Get-Content -LiteralPath $file).Count

    $lines = [System.Collections.Generic.Dictionary[int,object]]::new($lineCount)

    $fileMisses.ForEach({$lines[$_.Line] = 0})
    $fileHits.ForEach({$lines[$_.Line] = 1})
        
    for ($lineNum = 1; $lineNum -le $lineCount; $lineNum++) {
        if (-not $lines.ContainsKey($lineNum)) {
            $lines.Add($lineNum, $null)
        }
    }

    $lineReport = @($lines.GetEnumerator() | Sort-Object -Property Key | ForEach-Object -Process { $_.Value })

    $fileReport = New-Object -TypeName PSObject -Property @{name = "src/$fileName";source_digest=$digest.Hash;coverage=$lineReport}

    $fileReports.Add($fileReport)
}
$report = New-Object -TypeName PSOBject -Property @{
                service_name = $ServiceName
                service_job_id = $ServiceJobId
                repo_token = $RepoToken
                source_files = $fileReports
            }
$json = ConvertTo-Json -InputObject $report -Depth 3
$url = 'https://coveralls.io/api/v1/jobs'

# upload coverage report
Add-Type -AssemblyName System.Net.Http
$httpClient = [System.Net.Http.HttpClient]::new()
try {
    $content = [System.Net.Http.MultipartFormDataContent]::new()
    $fileContent = [System.Net.Http.StringContent]::new($json, [System.Text.Encoding]::UTF8, "application/json")
    $content.Add($fileContent, "json_file", "coverage-report.json");
    $response = $httpClient.PostAsync($url, $content).Result
} finally {
    $httpClient.Dispose()
}