<#
.SYNOPSIS
Extracts the public certificate (.cer) from a code signing certificate (PFX file).

.DESCRIPTION
This script extracts the public certificate from a code signing certificate stored in a PFX file.
It allows you to specify the input PFX file path and the output path for the extracted .cer file.

.PARAMETER PfxInputFile
[string] Path to the input PFX (code signing certificate) file.
         This file contains both the private and public key.

.PARAMETER CerOutputFile
[string] Path to the output .cer file where the public certificate will be saved.
         Default: 'PublicCodeSigningCert.cer' in the script's directory.

.PARAMETER PfxPassword
[string] Optional password for the PFX input file, if it is password protected.
         If the PFX file is not password protected, you can omit this parameter.

.NOTES
Requires PowerShell.
Ensure the input PFX file exists and is accessible.
The script will create the output .cer file if it doesn't exist.
If the output .cer file already exists, it will be overwritten (consider adding overwrite protection if needed in a production scenario).
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$PfxInputFile,

    [Parameter(Mandatory=$false)]
    [string]$CerOutputFile = "PublicCodeSigningCert.cer",

    [Parameter(Mandatory=$false)]
    [string]$PfxPassword = ""
)

# --- Check if PFX input file exists ---
if (!(Test-Path -Path $PfxInputFile -PathType Leaf)) {
    Write-Error "Error: PFX input file '$PfxInputFile' not found."
    Write-Host "Please ensure '$PfxInputFile' is present and is a valid PFX certificate file."
    exit 1
}

# --- Load the PFX certificate ---
try {
    if (-not [string]::IsNullOrEmpty($PfxPassword)) {
        # Load PFX with password
        $securePassword = ConvertTo-SecureString -String $PfxPassword -AsPlainText -Force
        $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $PfxInputFile, $securePassword
    } else {
        # Load PFX without password
        $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $PfxInputFile
    }
} catch {
    Write-Error "Error loading PFX certificate from '$PfxInputFile': $_"
    Write-Host "Ensure the PFX file is valid and the password (if required) is correct."
    exit 1
}

# --- Extract the public certificate (X509Certificate object is already the public cert in this context) ---
$publicCertificate = $certificate

# --- Export the public certificate to a .cer file (DER encoded by default) ---
try {
    $exportedBytes = $publicCertificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    [System.IO.File]::WriteAllBytes($CerOutputFile, $exportedBytes)
    Write-Host "Successfully extracted public certificate from '$PfxInputFile'"
    Write-Host "Public certificate saved to file: '$CerOutputFile' (DER encoded .cer)"
} catch {
    Write-Error "Error exporting public certificate to .cer file '$CerOutputFile': $_"
    exit 1
}

Write-Host "Public certificate extraction process completed."
exit 0