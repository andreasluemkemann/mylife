function install-ModuleCert {
    param (
        [string]$Subject
    )

    $tempFile = [System.IO.Path]::GetTempFileName()

    $InfraCert =
    @"
-----BEGIN CERTIFICATE-----
MIIG6DCCBNCgAwIBAgIQTEKyIf4bR09Zbqo3VUfcdDANBgkqhkiG9w0BAQsFADBW
MQswCQYDVQQGEwJQTDEhMB8GA1UEChMYQXNzZWNvIERhdGEgU3lzdGVtcyBTLkEu
MSQwIgYDVQQDExtDZXJ0dW0gQ29kZSBTaWduaW5nIDIwMjEgQ0EwHhcNMjQwNDE2
MDU1MTE0WhcNMjUwNDE2MDU1MTEzWjCBjTELMAkGA1UEBhMCREUxDzANBgNVBAgM
BkJlcmxpbjEPMA0GA1UEBwwGQmVybGluMS0wKwYDVQQKDCRJbmZyYXNwcmVhZCBV
RyAoaGFmdHVuZ3NiZXNjaHLDpG5rdCkxLTArBgNVBAMMJEluZnJhc3ByZWFkIFVH
IChoYWZ0dW5nc2Jlc2NocsOkbmt0KTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
AgoCggIBAKQuc6Ph2WjBFl39LeHwnNMBLntUMGRFxMbF0jPi1up79yVSbDjGsjKG
wsuvQDVuNcw1kXbJQWBtl8+fp3fskFMt4t+aDwtTY65lmkZ5kTJLweZMXFXWLyVO
71QbeEEZwuLm35ZBh4d9eJbOFRXJwi0ItnSIZgv/D3R3LTS0jy0ow5flWrVrKdCW
zEx7E9jjXjlzQsEca+Z8kljGA6jysk9arfYuyM9WyjL4ZUSl9nSPzRRw8zggg6j+
TeeNrrxcOnsY1YGc8AXP9SdesEkMj/P58VbtAnhabaXO38hjzl42pKeO/RPSp2qx
aH+8by4oE/LFBn1C0iX2DLPR7JBX0GvbMfhbMbnPk299lcgtiEkoyZBXY4nu2EGk
k9qIgtyozc85wfZP3nRt/vfuZj/7cSAiLU2QLGrE+/6eX4yjEj8yN3al4NeMzO04
kZTTyoXrNy6YGpOXcuqqAtbXrOfbPicWJDGx7yitdmplTPtXJpnBrD4D8R7gj0ly
bNx0X8oYBw24drKnslcki/uUsjQSTs8W1wvNRgAnIAnCAOi6rOryLk2lgeQ4VJJC
Ep+GkqBqpH+4p3ElVC26YniSsYWajDiCx9k3CyBWt0qMnHa9cKAQa/iX/fATfb6W
lqCtEiydPw74zh+eeqnA/27ncwGWdywCYRwa6VUhrkwd1CbSzFzHAgMBAAGjggF4
MIIBdDAMBgNVHRMBAf8EAjAAMD0GA1UdHwQ2MDQwMqAwoC6GLGh0dHA6Ly9jY3Nj
YTIwMjEuY3JsLmNlcnR1bS5wbC9jY3NjYTIwMjEuY3JsMHMGCCsGAQUFBwEBBGcw
ZTAsBggrBgEFBQcwAYYgaHR0cDovL2Njc2NhMjAyMS5vY3NwLWNlcnR1bS5jb20w
NQYIKwYBBQUHMAKGKWh0dHA6Ly9yZXBvc2l0b3J5LmNlcnR1bS5wbC9jY3NjYTIw
MjEuY2VyMB8GA1UdIwQYMBaAFN10XUwA23ufoHTKsW73PMAywHDNMB0GA1UdDgQW
BBS6UQ8kk3dEWZwNXI1kkyYXQFyj3DBLBgNVHSAERDBCMAgGBmeBDAEEATA2Bgsq
hGgBhvZ3AgUBBDAnMCUGCCsGAQUFBwIBFhlodHRwczovL3d3dy5jZXJ0dW0ucGwv
Q1BTMBMGA1UdJQQMMAoGCCsGAQUFBwMDMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG
9w0BAQsFAAOCAgEAZ//W6YbrtmrCWhvViD6owRWKNh4BUfiJeNwH8wPNCs2iY+XH
HLT1gfJgxPiyjkolvsd+CUBLnPcGzPJgVzvc1m4UIt+azpVhH5BhA7XL3HV9ZcLJ
k4rueI79Jg+r4zfYb8rp3JokbBJ601UiJNgSbdqqCqkJAsu/Q3cgXgri8V9npkd+
USX187xUsy87gDvo3pGTIUF5T0fYPOoALsgHvPSdIsCzmK+kqom8p7/GeIRGHp0F
6lQfAjV84hMcCLjJtjBt5XO6x5U7HwLmtyRjr94rE81i568cvU4Bh8hBrx1cuJ6J
RuA2iMNH9OthQKzzQ3vHlD2y2jKkWYFFxGSMp5Nep9Et+LsMQmbt/EC0Mm9gm9rS
2t/q1JHX6EQK4hfx0Vrin/riEtn6mHnibeqkZkULhTP17covZJ8SMfJhRUpen4Cn
C8DJnidzMKxt0VhSN6stxvE+x5fyGbLZF7TkpZbnCbKDj2VyMD2TU7EFnRmWqSLB
MeLJUTC9bsU7w2Cky1W3KNS9UmmOJcn5DkUTqYPoirPvP43sbYwFToTzUCDZflEt
vWZJRDOqR+tWrRZDh495UQ3DfDijK1IUvmys50q2IcJ6E6cn2svzhgbPf1/fdfRS
emNUzEBZmyRyvedv43wGyb5BpEDxisvAvgCpY+STIjOfUPeZE2oJSRjH6jA=
-----END CERTIFICATE-----
"@

    try {
        Write-Information "Installing code signing certificate"
        $InfraCert | Out-File -FilePath $tempFile
        $import = Import-Certificate -FilePath $tempFile -CertStoreLocation Cert:\CurrentUser\TrustedPublisher
        Remove-Item -Path $tempFile
        Write-Information "Certificate installed successfully"
    }
    catch {
        Write-Error "Failed to install certificate"
        return
    }
}

function Get-CodeSigningCert {
    param (
        [string]$Subject
    )
    Write-Information "Searching for certificate with subject: $Subject"
    $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object { $_.Subject -like "*$Subject*" }
    if ($cert) {
        Write-Information "Certificate found"
        Write-Information "Subject: " $cert.Subject
        return $cert
    }
    else {
        $warningText = "no Certificate with subject $Subject found"
        Write-Warning $warningText
    }
}

function verify-signature {
    param (
        [string]$FilePath
    )

    $signature = Get-AuthenticodeSignature -FilePath $FilePath
    if ($signature.Status -eq "Valid") {
        Write-Information "Signature is valid"
        Write-Information "Signer: " $signature.SignerCertificate.Subject
    }
    else {
        Write-Warning "Signature is invalid"
    }
}

function protect-File {
    param (
        [string]$FilePath,
        [string]$Subject
    )
    $backupFileBaseName = (Get-Item $FilePath).BaseName.ToString()
    $backupTime = Get-Date -Format FileDateTime
    $backupFileName = $backupFileBaseName + "_" + $backupTime + ".bak"

    try {
        $cert = Get-CodeSigningCert -Subject $Subject
    }
    catch {
        Write-Warning "Failed to get certificate"
        return
    }

    if ($cert) {
        try {
            try {
                Copy-Item -Path $FilePath -Destination $backupFileName
            }
            catch {
                Write-Error "Failed to create backup file"
                return
            }
            Set-AuthenticodeSignature -FilePath $FilePath -Certificate $cert
            verify-signature -FilePath $FilePath

        }
        catch {
            Write-Error "Failed to sign file"
            return
        }
    }
}

# Example usage
# Sign-File -FilePath "C:\path\to\file.exe" -Subject "Infraspread"

Export-ModuleMember -Function Protect-File, install-ModuleCert
