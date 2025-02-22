<#
.SYNOPSIS
Generates a password-less self-signed code signing certificate in PowerShell with customizable subject and organization.

.DESCRIPTION
This script generates a self-signed code signing certificate (.pfx) without password protection.
It allows you to customize the certificate's Subject (Common Name, Organization, Location, State, Country)
and other properties like DNS Name and validity period.

WARNING: Password-less code signing certificates are INSECURE and should ONLY be used for TESTING.

.PARAMETER CertFilePath
Specifies the output file path for the generated password-less .pfx certificate.
Default is "CodeSigningCert.pfx".

.PARAMETER SubjectName
Specifies the Common Name (CN) for the certificate's Subject.
Default is "TestCodeSigningCert-NoPassword".

.PARAMETER Organization
Specifies the Organization (O) for the certificate's Subject.
Default is "Test Organization".

.PARAMETER Location
Specifies the Locality (L) / City for the certificate's Subject.
Optional.

.PARAMETER State
Specifies the State or Province (S) for the certificate's Subject.
Optional.

.PARAMETER Country
Specifies the Country or Region (C) for the certificate's Subject (e.g., "US", "GB").
Optional.

.PARAMETER DnsName
Specifies the DNS name for the certificate. Default is "localhost".

.PARAMETER YearsValid
Specifies the validity period of the certificate in years. Default is 10.

.NOTES
Run this script in PowerShell.
Password-less certificates are highly discouraged for production or public code signing.
#>
[CmdletBinding()]
param (
    [string]$CertFilePath = "CodeSigningCert.pfx",
    [string]$SubjectName = "TestCodeSigningCert-NoPassword",
    [string]$Organization = "Test Organization",
    [string]$Location = "",          # Optional Location (Locality/City)
    [string]$State = "",             # Optional State/Province
    [string]$Country = "",           # Optional Country/Region (e.g., "US")
    [string]$DnsName = "localhost",
    [int]$YearsValid = 10
)

# --- Check if certificate file already exists ---
if (Test-Path -Path $CertFilePath) {
    Write-Host "File '$CertFilePath' already exists. Exiting..."
    exit
}

# --- Construct Subject string with provided parameters ---
$subjectStringBuilder = New-Object -TypeName System.Text.StringBuilder
$subjectStringBuilder.Append("CN=$SubjectName")
if (-not [string]::IsNullOrEmpty($Organization)) {
    $subjectStringBuilder.Append(",O=$Organization")
}
if (-not [string]::IsNullOrEmpty($Location)) {
    $subjectStringBuilder.Append(",L=$Location")
}
if (-not [string]::IsNullOrEmpty($State)) {
    $subjectStringBuilder.Append(",S=$State")
}
if (-not [string]::IsNullOrEmpty($Country)) {
    $subjectStringBuilder.Append(",C=$Country")
}
$CertSubject = $subjectStringBuilder.ToString()

# --- Generate Self-Signed Code Signing Certificate ---
try {
    $privateKey = New-SelfSignedCertificate -CertStoreLocation Cert:\CurrentUser\My `
        -DnsName $DnsName `
        -Subject $CertSubject `
        -Type CodeSigningCert `
        -KeySpec Signature `
        -KeyExportPolicy Exportable `
        -NotAfter (Get-Date).AddYears($YearsValid)

    Write-Host "Self-signed code signing certificate generated successfully."
    Write-Host "Subject: $($privateKey.Subject)"
    Write-Host "Thumbprint: $($privateKey.Thumbprint)"

    # --- Export to PFX without Password ---
    try {
        # Use an empty SecureString object for -Password to create a password-less PFX
        Export-PfxCertificate -Cert $privateKey -FilePath $CertFilePath -Password ([System.Security.SecureString]::new())
        Write-Host "Certificate exported to '$CertFilePath' (password-less)."
    }
    catch {
        Write-Error "Error exporting certificate to PFX: $_"
        Write-Warning "Make sure you have the necessary permissions to export the certificate."
        exit 1
    }

    Write-Host "---"
    Write-Host "Successfully generated password-less certificate: '$CertFilePath'"
    Write-Warning "WARNING: This PFX file is PASSWORD-LESS and INSECURE. Use ONLY for testing."

} catch {
    Write-Error "Error generating self-signed certificate: $_"
    exit 1
}

exit 0