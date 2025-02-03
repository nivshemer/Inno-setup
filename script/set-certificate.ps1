# Improved PowerShell Script for Setting Certificate and Modifying Hosts File

# Ensure Execution Policy is set to RemoteSigned
$policy = Get-ExecutionPolicy
if ($policy -ne "RemoteSigned") {
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
    Write-Host "Execution policy set to RemoteSigned."
} else {
    Write-Host "Execution policy is already set to RemoteSigned."
}
# 
# Define variables
$FTSERVERNAME = "OTD-FTSVRDEV"
$hostsFilePath = "$env:windir\System32\drivers\etc\hosts"
$jsonFilePath = "C:\Program Files (x86)\Nivshemer\Enforcer\appsettings.json"

try {
    # Read and modify appsettings.json safely
    if (Test-Path $jsonFilePath) {
        $jsonContent = Get-Content $jsonFilePath -Raw | ConvertFrom-Json
        if ($jsonContent.MoTApi -notmatch "^api\.") {
            $jsonContent.MoTApi = "api." + $jsonContent.MoTApi
            $jsonContent | ConvertTo-Json -Depth 10 | Set-Content $jsonFilePath
            Write-Host "Updated appsettings.json successfully."
        } else {
            Write-Host "MoTApi already contains 'api.'. No update needed."
        }
    } else {
        Write-Host "appsettings.json not found. Skipping update."
    }
} catch {
    Write-Host "Error modifying appsettings.json: $_"
}

# Modify hosts file safely
try {
    $hostsContent = Get-Content $hostsFilePath
    if ($hostsContent -notmatch "\b$FTSERVERNAME\b") {
        "127.0.0.1`t$FTSERVERNAME" | Add-Content $hostsFilePath
        Write-Host "Added $FTSERVERNAME to hosts file."
    } else {
        Write-Host "$FTSERVERNAME already exists in hosts file. No changes made."
    }
} catch {
    Write-Host "Error modifying hosts file: $_"
}
