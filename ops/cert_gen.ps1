if (Test-Path "CodeSigningCert.pfx") {
    Write-Host "CodeSigningCert.pfx already exists. Exiting..."
    exit
}

$privateKey = New-SelfSignedCertificate -CertStoreLocation Cert:\CurrentUser\My -DnsName "localhost" -Type CodeSigningCert -KeySpec Signature -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears(10)
$privateKey | Export-PfxCertificate -FilePath "CodeSigningCert.pfx" -Password (ConvertTo-SecureString -String "password" -Force -AsPlainText)