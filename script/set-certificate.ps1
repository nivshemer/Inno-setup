$policy = Get-ExecutionPolicy
$FTSERVERNAME="OTD-FTSVRDEV"
if ($policy -ne "RemoteSigned")
{
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
    Write-Host "Execution policy set to RemoteSigned." 
}
else
{
    Write-Host "Execution policy is already set to RemoteSigned." 
}

$hostsFilePath = "$env:windir\System32\drivers\etc\hosts"

[string]$jsonFilePath = "C:\Program Files (x86)\NanoLock\Enforcer\appsettings.json"

try {

    # Insert "api" before the first dot only
    $MoTApi = "$MoTDNSAddress" -replace "^([^.]+)\.", '$1-api.'

    # Define the entry to add to the hosts file
    $entry1 = "$MoTIpAddress`t$MoTApi"
    $entry2 = "$MoTIpAddress`t$MoTDNSAddress"
    $entry3 = "$FTIpAddress`t$FTSERVERNAME"

    # Read the hosts file content
    $hostsFileContent = Get-Content -Path $hostsFilePath

    # Add or update the entry in the hosts file
    if (-not ($hostsFileContent -contains $entry1)) {
        Add-Content -Path $hostsFilePath -Value $entry1
    }


    Start-Sleep -Seconds 1.5

    # Add or update the original entry
    if (-not ($hostsFileContent -contains $entry2)) {
        Add-Content -Path $hostsFilePath -Value $entry2
    }

    Start-Sleep -Seconds 1.5

    # Add or update the original entry
    if (-not ($hostsFileContent -contains $entry3)) {
        Add-Content -Path $hostsFilePath -Value $entry3
    }

    Write-Output "Hosts file updated successfully."

    # Extract the part of the input after the first dot
    $newValue = "$MoTDNSAddress" -replace "^[^.]+\.", ""

    # Read the file content as a single string
    $jsonContent = Get-Content -Path $jsonFilePath -Raw

    # Replace the value of "BaseDomain" with the new value using a simple string replace

    $jsonContent = $jsonContent -replace "nanolocksecurity.nl", $newValue

    # Write the updated content back to the file
    Set-Content -Path $jsonFilePath -Value $jsonContent -Force

    Write-Output "The value of 'BaseDomain' has been updated to '$newValue'."

# Define the registry paths
$registryPaths = @(
    "HKLM:\SOFTWARE\Siemens\Automation\Openness\15.0\PublicAPI\15.0.0.0",
    "HKLM:\SOFTWARE\Siemens\Automation\Openness\15.1\PublicAPI\15.1.0.0",
    "HKLM:\SOFTWARE\Siemens\Automation\Openness\16.0\PublicAPI\16.0.0.0",
    "HKLM:\SOFTWARE\Siemens\Automation\Openness\17.0\PublicAPI\17.0.0.0",
    "HKLM:\SOFTWARE\Siemens\Automation\Openness\18.0\PublicAPI\18.0.0.0"
)

# Define the base copy destination
$destinationBase = "C:\Program Files (x86)\NanoLock\Enforcer\SiemensPipeServer"

# Function to copy a file based on the Siemens.Engineering and Siemens.Engineering.Hmi values in the registry
function Copy-SiemensFilesFromRegistry {
    param (
        [string]$version,
        [string]$dllPath
    )

    # Construct the destination directory
    $destinationPath = "$destinationBase\$version"

    # Check if the file in the registry value exists
    if (Test-Path $dllPath) {
        # Create the destination directory if it doesn't exist
        if (-not (Test-Path $destinationPath)) {
            New-Item -Path $destinationPath -ItemType Directory -Force
        }

        # Copy the file from the path in the registry
        Copy-Item -Path $dllPath -Destination $destinationPath -Force
        Write-Host "Copied file from $dllPath to $destinationPath"
    } else {
        Write-Host "File not found at $dllPath (from registry)"
    }
}

# Iterate over each registry path
foreach ($path in $registryPaths) {
    # Extract version number (e.g., 15.0, 15.1, etc.)
    $version = $path -replace '.*Openness\\(\d+\.\d+).*', '$1'

    # Convert version to folder format (e.g., V18 for version 18.0)
    $folderVersion = "V" + $version

    # Check if the registry path exists
    if (Test-Path $path) {
        Write-Host "Registry path exists: $path"

        # Get the value of Siemens.Engineering from the registry and copy the corresponding file
        $engineering = Get-ItemProperty -Path $path -Name "Siemens.Engineering" -ErrorAction SilentlyContinue
        if ($engineering) {
            $engineeringDllPath = $engineering.'Siemens.Engineering'
            Copy-SiemensFilesFromRegistry -version $folderVersion -dllPath $engineeringDllPath
        } else {
            Write-Host "Siemens.Engineering value not found in $path"
        }

        # Get the value of Siemens.Engineering.Hmi from the registry and copy the corresponding file
        $hmi = Get-ItemProperty -Path $path -Name "Siemens.Engineering.Hmi" -ErrorAction SilentlyContinue
        if ($hmi) {
            $hmiDllPath = $hmi.'Siemens.Engineering.Hmi'
            Copy-SiemensFilesFromRegistry -version $folderVersion -dllPath $hmiDllPath
        } else {
            Write-Host "Siemens.Engineering.Hmi value not found in $path"
        }

    } else {
        Write-Host "Registry path does not exist: $path"
    }

    Write-Host "-----------------------------"
}


# Get all directories under the target path
$directories = Get-ChildItem -Path $destinationBase -Directory

# Iterate through each directory
foreach ($dir in $directories) {
    # Check if the directory name contains ".0"
    if ($dir.Name -like "*.0*") {
        # Remove ".0" from the directory name
        $newName = $dir.Name -replace "\.0", ""

        # Construct the full path for the new directory name
        $newPath = Join-Path $dir.Parent.FullName $newName

        # Check if the target directory already exists
        if (Test-Path $newPath) {
            Write-Host "Directory '$newPath' already exists. Merging files..."

            # Merge files from the ".0" directory into the existing directory
            Get-ChildItem -Path $dir.FullName -Recurse | ForEach-Object {
                $destination = Join-Path $newPath $_.Name

                # If the file exists, it will be overwritten
                if (Test-Path $destination) {
                    Write-Host "File '$destination' already exists. Overwriting..."
                }

                # Copy the file or directory to the new location
                Copy-Item -Path $_.FullName -Destination $newPath -Recurse -Force
            }

            # Remove the old ".0" directory after merging
            Remove-Item -Path $dir.FullName -Recurse -Force
            Write-Host "Removed old directory '$($dir.FullName)'"
        } else {
            # Rename the directory if the target does not exist
            Rename-Item -Path $dir.FullName -NewName $newPath
            Write-Host "Renamed '$($dir.FullName)' to '$newPath'"
        }
    }
}


# Define variables
$website = "$MoTDNSAddress"
$certificateFilePath = "C:\temp\cert.cer"

# Ensure temporary directory exists
if (-not (Test-Path "C:\temp")) {
    New-Item -Path "C:\" -Name "temp" -ItemType "Directory" | Out-Null
}

# Function to export a certificate from a website
function Export-CertificateFromWebsite {
    param (
        [string]$HostName,
        [string]$OutputPath
    )

    try {
        # Bypass SSL validation to extract the certificate
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }

        # Establish an SSL connection to retrieve the certificate
        $tcpClient = New-Object Net.Sockets.TcpClient($HostName, 443)
        $sslStream = New-Object System.Net.Security.SslStream($tcpClient.GetStream(), $false, ({ $true }))

        # Authenticate as a client to retrieve server certificate
        $sslStream.AuthenticateAsClient($HostName)
        $certificate = $sslStream.RemoteCertificate
        $tcpClient.Close()

        # Save the certificate in DER format to the specified path
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certificate)
        [System.IO.File]::WriteAllBytes($OutputPath, $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
        Write-Host "Certificate successfully exported to $OutputPath" -ForegroundColor Green
    } catch {
        Write-Host "Failed to export certificate: $_" -ForegroundColor Red
        throw $_
    }
}

# Function to install a certificate into the Trusted Root store
function Install-CertificateToTrustedRoot {
    param (
        [string]$CertPath
    )

    try {
        $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertPath)
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
        $store.Open("ReadWrite")
        $store.Add($certificate)
        $store.Close()
        Write-Host "Certificate successfully installed to Trusted Root Certification Authorities" -ForegroundColor Green
    } catch {
        Write-Host "Failed to install certificate: $_" -ForegroundColor Red
        throw $_
    }
}

# Execute functions
Export-CertificateFromWebsite -HostName $website -OutputPath $certificateFilePath
Install-CertificateToTrustedRoot -CertPath $certificateFilePath



}
catch
{
    Write-Host "Error occurred during host entries registration: $_" 
}
