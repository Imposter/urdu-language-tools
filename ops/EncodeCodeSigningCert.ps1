<#
.SYNOPSIS
Encodes a PFX certificate file to a Base64 string, with default input/output files and overwrite protection.

.DESCRIPTION
This script encodes a PFX certificate file to a Base64 string.

By default, it reads the PFX certificate from 'CodeSigningCert.pfx' (if it exists in the same directory)
and saves the Base64 encoded string to 'EncodedCert.base64' (in the same directory).

You can optionally specify different input and output file paths using the -InputFile and -OutputFile parameters.

The script will **fail and exit** if the output file 'EncodedCert.base64' (or the file specified by -OutputFile) already exists, to prevent accidental overwriting.

.PARAMETER InputFile
[string] Path to the PFX certificate file to encode.
         Default: 'CodeSigningCert.pfx' in the script's directory.

.PARAMETER OutputFile
[string] Path to the file where the Base64 encoded string will be saved.
         Default: 'EncodedCert.base64' in the script's directory.

.NOTES
Requires PowerShell.
Ensure the input PFX file exists and is accessible.
The script will create the output file if it doesn't exist.
If the output file already exists, the script will output an error and terminate.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$InputFile = "CodeSigningCert.pfx",  # Default input PFX file name

    [Parameter(Mandatory=$false)]
    [string]$OutputFile = "EncodedCert.base64"    # Default output Base64 file name
)

# --- Check if output Base64 file already exists ---
if (Test-Path -Path $OutputFile -PathType Leaf) {
    Write-Error "Error: Output Base64 file '$OutputFile' already exists."
    Write-Host "To prevent accidental overwriting, the script will now exit."
    exit 1
}

# --- Check if PFX input file exists ---
if (!(Test-Path -Path $InputFile -PathType Leaf)) {
    Write-Error "Error: PFX input file '$InputFile' not found."
    Write-Host "Please ensure '$InputFile' is present and is a valid PFX certificate file."
    exit 1
}

# --- Read PFX file as bytes ---
try {
    $certificateBytes = [System.IO.File]::ReadAllBytes($InputFile)
} catch {
    Write-Error "Error reading PFX file '$InputFile': $_"
    exit 1
}

# --- Base64 encode the byte array ---
try {
    $base64String = [System.Convert]::ToBase64String($certificateBytes)
} catch {
    Write-Error "Error encoding to Base64: $_"
    exit 1
}

# --- Save Base64 content to file ---
try {
    $base64String | Out-File -FilePath $OutputFile -Encoding utf8
    Write-Host "Successfully encoded PFX from '$InputFile' and saved Base64 content to file: '$OutputFile'"
} catch {
    Write-Error "Error saving Base64 content to file '$OutputFile': $_"
    exit 1
}

# --- Output to console ---
Write-Host "Base64 Encoded Content:"
Write-Host $base64String

Write-Host "Encoding process completed."
exit 0