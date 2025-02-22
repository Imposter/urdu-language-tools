# --- Configuration ---
$PfxFilePath = "CodeSigningCert.pfx"  # Path to your CodeSigningCert.pfx file
$OutputFile = "EncodedCert.txt"       # Optional: File to save the Base64 output

# --- Check if PFX file exists ---
if (!(Test-Path -Path $PfxFilePath -PathType Leaf)) {
    Write-Error "Error: PFX file '$PfxFilePath' not found."
    exit 1
}

# --- Read PFX file as bytes ---
try {
    $certificateBytes = [System.IO.File]::ReadAllBytes($PfxFilePath)
} catch {
    Write-Error "Error reading PFX file '$PfxFilePath': $_"
    exit 1
}

# --- Base64 encode the byte array ---
try {
    $base64String = [System.Convert]::ToBase64String($certificateBytes)
} catch {
    Write-Error "Error encoding to Base64: $_"
    exit 1
}

# --- Output to console ---
Write-Host "Base64 Encoded Content:"
Write-Host $base64String

# --- Optional: Save to file ---
if (-not [string]::IsNullOrEmpty($OutputFile)) {
    try {
        $base64String | Out-File -FilePath $OutputFile -Encoding utf8
        Write-Host "Base64 content saved to file: '$OutputFile'"
    } catch {
        Write-Warning "Warning: Error saving Base64 content to file '$OutputFile': $_"
    }
}