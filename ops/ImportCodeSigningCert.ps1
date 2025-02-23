<#
.SYNOPSIS
Imports a code signing certificate (.pfx) into the Windows certificate store.

.DESCRIPTION
This script imports a code signing certificate from a Base64 encoded string or a .pfx file into the specified Windows certificate store location.
It supports both password-protected and password-less PFX files.

If a Base64 encoded certificate string is provided via the -CertificateBase64 parameter, the script will first decode it and save it to a .pfx file specified by -CertificateFilePath.
If no -CertificateBase64 is provided, the script will attempt to import an existing .pfx file from the path specified by -CertificateFilePath.

.PARAMETER CertificateBase64
[string] Optional. A Base64 encoded string of the code signing certificate (.pfx).
         If provided, the script will decode this string and save it to the file specified by -CertificateFilePath.
         If not provided, the script assumes the .pfx file already exists at the path specified by -CertificateFilePath.

.PARAMETER CertificatePassword
[string] Optional. The password for the code signing certificate (.pfx) if it is password protected.
         If the PFX file is password-less, you can omit this parameter.

.PARAMETER CertificateFilePath
[string] Optional. The file path where the .pfx certificate will be saved (if -CertificateBase64 is provided) or the path to an existing .pfx file to import.
         Default: "UrduLanguageTools/CodeSigningCert.pfx"

.PARAMETER CertStoreLocation
[string] Optional. The Windows certificate store location where the certificate will be imported.
         Default: "Cert:\CurrentUser\My" (User's personal certificate store)

.NOTES
This script requires PowerShell to run.
It is designed to be used in automated environments like GitHub Actions or local PowerShell sessions.

.EXAMPLE
# Example 1: Import from Base64 string with password, to the default certificate store
Import-CodeSigningCert.ps1 -CertificateBase64 "<Base64 String Here>" -CertificatePassword "YourPassword"

.EXAMPLE
# Example 2: Import from Base64 string without password, to the default certificate store
Import-CodeSigningCert.ps1 -CertificateBase64 "<Base64 String Here>"

.EXAMPLE
# Example 3: Import an existing PFX file (password-less) to the LocalMachine certificate store
Import-CodeSigningCert.ps1 -CertificateFilePath "C:\Certs\ExistingCert.pfx" -CertStoreLocation "Cert:\LocalMachine\My"

.EXAMPLE
# Example 4: Import an existing PFX file (password-protected) to the CurrentUser certificate store
Import-CodeSigningCert.ps1 -CertificateFilePath "C:\Certs\SecureCert.pfx" -CertificatePassword "SecurePassword" -CertStoreLocation "Cert:\CurrentUser\My"
#>
param (
    [Parameter(Mandatory=$false)]
    [string]$CertificateBase64,

    [Parameter(Mandatory=$false)]
    [string]$CertificatePassword,

    [Parameter(Mandatory=$false)]
    [string]$CertificateFilePath = "UrduLanguageTools/CodeSigningCert.pfx", # More specific default path

    [Parameter(Mandatory=$false)]
    [string]$CertStoreLocation = "Cert:\CurrentUser\My"
)

# If no base64 string is provided, then check if the CertificateFilePath exists
if ([string]::IsNullOrEmpty($CertificateBase64)) {
    if (-not (Test-Path -Path $CertificateFilePath -PathType Leaf)) {
        Write-Error "Error: Certificate file '$CertificateFilePath' not found."
        Write-Host "Please ensure '$CertificateFilePath' is present and is a valid PFX certificate file or provide a valid base64 string."
        exit 1
    }
} else {
    # --- Save Code Signing Certificate from Base64 ---
    try {
        $certificateBytes = [System.Convert]::FromBase64String($CertificateBase64)
        [System.IO.File]::WriteAllBytes($CertificateFilePath, $certificateBytes)
        Write-Host "Code signing certificate saved to '$CertificateFilePath'"
    } catch {
        Write-Error "Error saving code signing certificate to file: $_"
        exit 1
    }
}

# --- Import Code Signing Certificate to Certificate Store ---
try {
    if (-not [string]::IsNullOrEmpty($CertificatePassword)) {
        # Import with Password if provided
        $securePassword = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
        Import-PfxCertificate -FilePath $CertificateFilePath `
        -Password $securePassword `
        -CertStoreLocation $CertStoreLocation
    } else {
        # Import without Password if $CertificatePassword is empty
        Import-PfxCertificate -FilePath $CertificateFilePath `
        -Password ([System.Security.SecureString]::new()) `
        -CertStoreLocation $CertStoreLocation
        Write-Warning "Importing PFX without a password. This is generally discouraged for security reasons unless the PFX is intentionally password-less."
    }
    Write-Host "Code signing certificate imported to certificate store: '$CertStoreLocation'"
} catch {
    Write-Error "Error importing code signing certificate to certificate store: $_"
    exit 1
}

Write-Host "Code signing certificate import process completed."
exit 0