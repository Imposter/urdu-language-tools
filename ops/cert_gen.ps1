<#
.SYNOPSIS
Generates a self-signed code signing certificate in PowerShell with customizable subject, organization, and optional password protection.

.DESCRIPTION
This script generates a self-signed code signing certificate (.pfx).
It allows you to customize the certificate's Subject (Common Name, Organization, Location, State, Country)
and other properties like DNS Name and validity period.
You can also optionally protect the generated PFX file with a password.

.PARAMETER CertFilePath
Specifies the output file path for the generated .pfx certificate.
Default is "CodeSigningCert.pfx".

.PARAMETER SubjectName
Specifies the Common Name (CN) for the certificate's Subject.
Default is "TestCodeSigningCert".

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

.PARAMETER PfxPassword
[string] Optional password to protect the exported .pfx certificate.
         If not provided, the PFX file will be generated without a password (INSECURE, for TESTING ONLY).

.NOTES
Run this script in PowerShell.
Password-less certificates are highly discouraged for production or public code signing.
For production, ALWAYS use a strong password to protect your code signing certificate.
#>
[CmdletBinding()]
param (
    [string]$CertFilePath = "CodeSigningCert.pfx",
    [string]$SubjectName = "TestCodeSigningCert",
    [string]$Organization = "Test Organization",
    [string]$Location = "",          # Optional Location (Locality/City)
    [string]$State = "",             # Optional State/Province
    [string]$Country = "",           # Optional Country/Region (e.g., "US")
    [string]$DnsName = "localhost",
    [int]$YearsValid = 10,
    [string]$PfxPassword = ""       # Optional password for PFX file
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

    # --- Export to PFX with or without Password based on $PfxPassword ---
    try {
        if (-not [string]::IsNullOrEmpty($PfxPassword)) {
            # Export to PFX with Password
            $securePassword = ConvertTo-SecureString -String $PfxPassword -AsPlainText -Force
            Export-PfxCertificate -Cert $privateKey -FilePath $CertFilePath -Password $securePassword
            Write-Host "Certificate exported to '$CertFilePath' (password-protected)."
        } else {
            # Export to PFX without Password (password-less)
            Export-PfxCertificate -Cert $privateKey -FilePath $CertFilePath -Password ([System.Security.SecureString]::new())
            Write-Host "Certificate exported to '$CertFilePath' (password-less)."
            Write-Warning "WARNING: This PFX file is PASSWORD-LESS and INSECURE. Use ONLY for testing."
        }
    }
    catch {
        Write-Error "Error exporting certificate to PFX: $_"
        Write-Warning "Make sure you have the necessary permissions to export the certificate."
        exit 1
    }

    Write-Host "---"
    if (-not [string]::IsNullOrEmpty($PfxPassword)) {
        Write-Host "Successfully generated password-protected certificate: '$CertFilePath'"
    } else {
        Write-Host "Successfully generated password-less certificate: '$CertFilePath'"
    }


} catch {
    Write-Error "Error generating self-signed certificate: $_"
    exit 1
}

exit 0