﻿$modules = @("Pester", "PSFramework", "PSModuleDevelopment", "PSScriptAnalyzer","AzureAD")

foreach ($module in $modules) {
    Write-Host "Installing $module" -ForegroundColor Cyan
    Install-Module $module -Force -SkipPublisherCheck
    Import-Module $module -Force -PassThru
}