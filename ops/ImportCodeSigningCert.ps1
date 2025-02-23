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