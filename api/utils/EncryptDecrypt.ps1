# Usage
# .\EncryptDecrypt.ps1 -Mode decrypt -SecretString "HDgYdLdHLY8dhGDjd..." -Text "bhzhmBIMaYbKAsrIcfVb07X_3Lh0pyUFjCIBNSOmxLs"
#

param(
    [Parameter(Mandatory)]
    [ValidateSet("encrypt","decrypt")]
    [string]$Mode,

    [Parameter(Mandatory)]
    [string]$SecretString,   # The EXACT raw string stored in Key Vault — NOT decoded, NOT modified

    [Parameter(Mandatory)]
    [string]$Text            # Plain text (encrypt mode) or encrypted text (decrypt mode)
)

# -------------------------------------------------------------------------------------
# This function mirrors the C# line:
#     var key = SHA256.HashData(UTF8(secretString));
#
# IMPORTANT:
# - The secret is treated as a normal string
# - No Base64 decoding is performed (even if the secret LOOKS like Base64)
# - Key is ALWAYS SHA256(UTF8(secretString)) → 32 bytes
#
# -------------------------------------------------------------------------------------
function Get-AesKeyFromSecretString($secretString) {
    $sha = [System.Security.Cryptography.SHA256]::Create()

    # Convert secret to UTF-8 bytes (C# default encoding)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($secretString)

    # Generate the 32-byte AES key
    return $sha.ComputeHash($bytes)
}

# -------------------------------------------------------------------------------------
# Encrypt using:
#   - AES-256
#   - CBC mode
#   - PKCS7 padding
#   - FIXED IV (same as C#)
#   - URL-safe Base64 output
#
# IMPORTANT: deterministic encryption (same plaintext → same ciphertext)
# -------------------------------------------------------------------------------------
function Encrypt-Deterministic($plainText, $keyBytes) {

    # Create AES instance configured EXACTLY like the C# code
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.Mode    = 'CBC'
    $aes.Padding = 'PKCS7'
    $aes.Key     = $keyBytes

    # 16-byte fixed IV — MUST match C# exactly
    $aes.IV = [System.Text.Encoding]::UTF8.GetBytes("16bytes-fixed-iv")

    $encryptor = $aes.CreateEncryptor()

    # Convert plaintext to UTF-8 bytes
    $plainBytes = [System.Text.Encoding]::UTF8.GetBytes($plainText)

    # Perform AES encryption
    $encrypted  = $encryptor.TransformFinalBlock($plainBytes, 0, $plainBytes.Length)

    # Convert to Base64, then to URL-safe format (matches your C# helper)
    return [Convert]::ToBase64String($encrypted).
        Replace("+","-").
        Replace("/","_").
        TrimEnd("=")
}

# -------------------------------------------------------------------------------------
# Decrypt using the same AES settings and fixed IV.
# Converts URL-safe Base64 back to normal Base64 before decoding.
# -------------------------------------------------------------------------------------
function Decrypt-Deterministic($cipherText, $keyBytes) {
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.Mode    = 'CBC'
    $aes.Padding = 'PKCS7'
    $aes.Key     = $keyBytes
    $aes.IV      = [System.Text.Encoding]::UTF8.GetBytes("16bytes-fixed-iv")

    # Convert URL-safe Base64 back to standard Base64
    $b64 = $cipherText.Replace("-","+").Replace("_","/")
    switch ($b64.Length % 4) {
        2 { $b64 += "==" }
        3 { $b64 += "=" }
    }

    # Decode Base64 → encrypted bytes
    $encBytes = [Convert]::FromBase64String($b64)

    # Decrypt using the AES key + IV
    $decryptor = $aes.CreateDecryptor()
    $plainBytes = $decryptor.TransformFinalBlock($encBytes, 0, $encBytes.Length)

    # Convert back to UTF-8 string
    return [System.Text.Encoding]::UTF8.GetString($plainBytes)
}

# -------------------------------------------------------------------------------------
# Generate the AES key BY HASHING the key vault secret
# This step MUST be identical to the C#:
#   SHA256(UTF8(secretString))
# -------------------------------------------------------------------------------------
$keyBytes = Get-AesKeyFromSecretString $SecretString

# -------------------------------------------------------------------------------------
# Execute encryption or decryption
# -------------------------------------------------------------------------------------
switch ($Mode) {
    "encrypt" { Encrypt-Deterministic $Text $keyBytes }
    "decrypt" { Decrypt-Deterministic $Text $keyBytes }
}
