function Read-File {
    param (
        [string]$fileName
    )
    Get-Content -Path $fileName
}

function ParseJson {
    param (
        [string]$plainText
    )
    $json = ConvertFrom-Json -InputObject $plainText
    return $json
}

function Read-ConfigFile {
    $configTxt = Read-File -fileName "config.json"
    $configJson = ParseJson -plainText $configTxt
    return $configJson
}
