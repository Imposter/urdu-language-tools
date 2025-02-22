<#
.SYNOPSIS
Decodes a Base64 encoded string to a PFX certificate file, with default input/output files and overwrite protection.

.DESCRIPTION
This script decodes a Base64 encoded string and saves it as a PFX certificate file.

By default, it reads the Base64 string from 'EncodedCert.base64' (if it exists in the same directory)
and saves the decoded PFX certificate to 'CodeSigningCert.pfx' (in the same directory).

You can optionally specify different input and output file paths using the -InputFile and -OutputPfxFile parameters.

The script will **fail and exit** if the output file 'CodeSigningCert.pfx' (or the file specified by -OutputPfxFile) already exists, to prevent accidental overwriting.

.PARAMETER InputFile
[string] Path to the file containing the Base64 encoded string.
         Default: 'EncodedCert.base64' in the script's directory.

.PARAMETER OutputPfxFile
[string] Path to the file where the decoded PFX certificate will be saved.
         Default: 'CodeSigningCert.pfx' in the script's directory.

.NOTES
Requires PowerShell.
Ensure the input file contains valid Base64 encoded certificate data.
The script will create the output PFX file if it doesn't exist.
If the output PFX file already exists, the script will output an error and terminate.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$InputFile = "EncodedCert.base64",  # Default input file name

    [Parameter(Mandatory=$false)]
    [string]$OutputPfxFile = "CodeSigningCert.pfx" # Default output file name
)

# --- Check if output PFX file already exists ---
if (Test-Path -Path $OutputPfxFile -PathType Leaf) {
    Write-Error "Error: Output PFX file '$OutputPfxFile' already exists."
    Write-Host "To prevent accidental overwriting, the script will now exit."
    exit 1
}

# --- Check if input Base64 file exists ---
if (!(Test-Path -Path $InputFile -PathType Leaf)) {
    Write-Error "Error: Input Base64 file '$InputFile' not found."
    Write-Host "Please ensure '$InputFile' is present and contains the Base64 encoded certificate data."
    exit 1
}

# --- Read Base64 string from input file ---
try {
    $base64String = Get-Content -Path $InputFile -Raw
} catch {
    Write-Error "Error reading input file '$InputFile': $_"
    exit 1
}

# --- Base64 decode to byte array ---
try {
    $certificateBytes = [System.Convert]::FromBase64String($base64String)
} catch {
    Write-Error "Error decoding from Base64: $_"
    Write-Host "Make sure the content of '$InputFile' is a valid Base64 encoded string."
    exit 1
}

# --- Write byte array to PFX file ---
try {
    [System.IO.File]::WriteAllBytes($OutputPfxFile, $certificateBytes)
    Write-Host "Successfully decoded Base64 from '$InputFile' and saved PFX file to: '$OutputPfxFile'"
} catch {
    Write-Error "Error writing PFX file '$OutputPfxFile': $_"
    exit 1
}

Write-Host "Decoding process completed."
exit 0