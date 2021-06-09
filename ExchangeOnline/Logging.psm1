$global:LogFolderName = $null
$global:LogFileName = $null
$global:Log = $null

$global:LogFolderName = "logs\"
$global:LogFileName = $LogFolderName
$global:LogFileName += Get-Date -Format o | foreach {$_ -replace ":", "."}
$global:LogFileName += ".log"

function SetUpLogs {
    $global:Log = $null

    if (Test-Path $LogFolderName) {
        Write-Host "Log folder exists"
    } else {
        mkdir $LogFolderName
    }
}

function WriteLog {
    Param ([string]$_logstring)

    $global:Log += $_logstring
    $global:Log += "`r`n"
}

function WriteLogToFile {
    Param(
        [string]$_logFile,
        [string]$_logString
    )

    Add-content $_logFile -value $_logString
}
