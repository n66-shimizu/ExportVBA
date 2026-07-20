#
# 共通関数
# ---------------------------------------------------
# ExportVBA共通関数
# 2026.06.20  PrimeBrains 清水
#

###############################
## 共通関数
###############################

##################################
## フォルダがない場合に作成
## 存在する場合はフォルダ内をクリア
##################################

function InitFolder {
    param (
        [string]$folderPath
    )

    if (-Not (Test-Path -Path $folderPath)) {
        LogDebug ("Create:" + $folderPath)
        New-Item -Path "$folderPath" -ItemType Directory    
    }
}

##################################
## フォルダがない場合に作成
## 存在する場合はフォルダ内をクリア
##################################

function InitFolderWithClear {
    param (
        [string]$folderPath
    )

    if (Test-Path -Path $folderPath) {
        LogDebug ("Clear:" + $folderPath)
        Get-ChildItem $folderPath -Force | Remove-Item -Recurse -Force
    } else {
        LogDebug ("Create:" + $folderPath)
        New-Item -Path "$folderPath" -ItemType Directory    
    }
}



##################################
## datetimeの出力形式を編集
##################################

function ReformDatetime {
    param (
        [datetime]$datetime
    )

    if ($datetime -ne $null) {
        return [datetime]::Parse($datetime).ToString("yyyy/MM/dd HH:mm:ss")
    } else {
        return ""
    }
}

##################################
## エクスポート親フォルダ取得
##################################

function GetExportFolder {

    param (
        [string]$targetFile,
        [string]$exportFolder
    )
    
    LogDebug ("GetExportPath targetFile=" + $targetFile)
    LogDebug ("GetExportPath exportFolder=" + $exportFolder)
    
    if ($exportFolder -eq "") {
        # exportフォルダが未指定の場合は、親フォルダ配下のEXPORTフォルダを使用
        $exportFolder =  Join-Path -Path ( Split-Path -Path $targetFile -Parent )  -ChildPath "ExportedVBA"
    }
    return $exportFolder
}

##################################
## エクスポートフォルダ取得
##################################

function GetExportPath {

    param (
        [string]$targetFile,
        [string]$exportFolder
    )
    
    LogDebug ("GetExportPath targetFile=" + $targetFile)
    LogDebug ("GetExportPath exportFolder=" + $exportFolder)
    
    $leaf = [System.IO.Path]::GetFileNameWithoutExtension($targetFile)   # targetファイルから拡張子以外の部分を取り出す
    $extension = [System.IO.Path]::GetExtension($targetFile)             # targetファイルから拡張子を取り出す
    $extension = $extension -replace '\.', '_'                           # 拡張子の'.'を'_'に変換
    $exportPath = Join-Path -Path $exportFolder -ChildPath ( $leaf + $extension )
    
    LogDebug ("GetExportPath leaf=" + $leaf)
    LogDebug ("GetExportPath extension=" + $extension)
    LogDebug ("GetExportPath exportPath=" + $exportPath)

    return $exportPath
}

###############################
## ログ初期化
###############################

function InitLog {

    param (
        [int]$logLevel
    )
    $script:logLevel = $logLevel

    #log初期化
    $script:log = ''
    LogInfo ""

}

###############################
## ログ終了
###############################

function EndLog {

    LogInfo ""
}


###############################
## ログ（エラー）
###############################

function LogError
{
    param (
        [string]$message
    )
    _logFormat -level 'ERROR', -message $message
}

###############################
## ログ（インフォ）
###############################

function LogInfo
{
    param (
        [string]$message
    )
    if ($script:logLevel -lt 2) {
        return
    }
    _logFormat -level 'INFO ' -message $message
}

###############################
## ログ（デバック）
###############################

function LogDebug
{
    param (
        [string]$message
    )
    if ($script:logLevel -lt 9) {
        return
    }
    _logFormat -level 'DEBUG' -message $message
}

############################################
## ログ編集出力
## LogError,LogInfo,LogDebugから呼ばれる前提
#############################################

function _LogFormat
{
    param (
        [string]$level,
        [string]$message
    )
    $func = (Get-PSCallStack)[2].functionName
    $time = Get-Date -Format 'yy/MM/dd HH:mm:ss'
    Write-Host ( $time + ' ' + $level + ' [' + $func + '] ' + $message ) 
}

