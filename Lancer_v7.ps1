$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
if (-not $env:MINFI_USERNAME) {
    $env:MINFI_USERNAME = Read-Host 'Identifiant de votre session locale'
}
if (-not $env:MINFI_PASSWORD) {
    $secretV7 = Read-Host 'Choisissez le mot de passe de votre session locale' -AsSecureString
    $ptrV7 = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secretV7)
    try {
        $env:MINFI_PASSWORD = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptrV7)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptrV7)
    }
}
python -m streamlit run app7.py --server.address 127.0.0.1 --server.port 8507
