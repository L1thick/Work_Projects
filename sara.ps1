﻿# Powershell Script to download SaRACmd, extract, remove Office.
# This script does not care which version it is, it will remove Office;
# v3.3

# NOTES:
# Intune install command: Powershell.exe -executionpolicy bypass -File sara.ps1
# Detection Rule for Intune: IF SOFTWARE VERSION is less than "16.0.17830.20138", require software.

# SaRACMD URL. Change if needed:
$url = "https://aka.ms/SaRA_CommandLineVersionFiles"

# Create working directory & manage the file:
New-Item -Path 'C:\saratemp\' -ItemType Directory
Invoke-WebRequest -Uri $url -OutFile "C:\saratemp\download.zip"
Expand-Archive -LiteralPath 'C:\saratemp\download.zip' -DestinationPath 'C:\saratemp\expanded\'

# Close open Office applications:
$officeProcesses = "lync","winword","excel","msaccess","mstore","infopath","setlang","msouc","ois","onenote","outlook","powerpnt","mspub","groove","visio","winproj","graph","teams"
foreach ($process in $officeProcesses) {
    try {
        Stop-Process -Name $process -Force -ErrorAction Stop
    } catch {
        Write-Host "Process was not running: $process"
    }
}

# Remove Office versions:
& "C:\saratemp\expanded\SaRACmd.exe" -S OfficeScrubScenario -AcceptEula -OfficeVersion All

# Cleanup after yourself:
Remove-Item -LiteralPath "C:\saratemp\" -Force -Recurse
Write-Host "Deployment removed. Process complete." 
