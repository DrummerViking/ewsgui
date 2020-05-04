$modules = @("Pester", "PSFramework", "PSModuleDevelopment", "PSScriptAnalyzer", "ExchangeOnlineManagement")

foreach ($module in $modules) {
    Write-Host "Installing $module" -ForegroundColor Cyan
    Install-Module $module -Force -SkipPublisherCheck
    Import-Module $module -Force -PassThru
}