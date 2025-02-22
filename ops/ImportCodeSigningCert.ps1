param (
    [Parameter(Mandatory=$true)]
    [string]$CertificateBase64,

    [Parameter(Mandatory=$false)]
    [string]$CertificatePassword,

    [Parameter(Mandatory=$false)]
    [string]$CertificateFilePath = "UrduLanguageTools/CodeSigningCert.pfx", # More specific default path

    [Parameter(Mandatory=$false)]
    [string]$CertStoreLocation = "Cert:\CurrentUser\My"
)

# --- Save Code Signing Certificate from Base64 ---
try {
    $certificateBytes = [System.Convert]::FromBase64String($CertificateBase64)
    [System.IO.File]::WriteAllBytes($CertificateFilePath, $certificateBytes)
    Write-Host "Code signing certificate saved to '$CertificateFilePath'"
} catch {
    Write-Error "Error saving code signing certificate to file: $_"
    exit 1
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