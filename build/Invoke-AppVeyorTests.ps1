Import-Module -Name Pester -MinimumVersion 5.2.2
$res = Invoke-Pester -Path "$PSSCRIPTROOT/../test" -OutputFormat NUnitXml -OutputFile TestsResults.xml -PassThru -CodeCoverage 'src\\*-XL*.ps1'

(New-Object 'System.Net.WebClient').UploadFile("https://ci.appveyor.com/api/testresults/nunit/$($env:APPVEYOR_JOB_ID)", (Resolve-Path .\TestsResults.xml))

& "$PSSCRIPTROOT/Publish-CoverageReport.ps1" -RepoToken $env:CoverallsRepoToken -CodeCoverage $res.CodeCoverage -ServiceName 'appveyor' -ServiceJobId $env:APPVEYOR_JOB_ID -GitBranch $env:APPVEYOR_REPO_BRANCH

if ($res.FailedCount -gt 0) { throw "$($res.FailedCount) tests failed."}
