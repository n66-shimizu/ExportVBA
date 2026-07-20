#
# 別ウィンドウ表示用ログ関数
# ---------------------------------------------------
# ExportVBAforSVN共通関数
# Common.ps1をオーバーライド
# 2026.06.20  PrimeBrains 清水
#

###############################
## 共通ログ関数（別ウィンドウ）
###############################

###############################
## ログ初期化（別ウィンドウ）
###############################

function initWindowLog {

    param (
        [int]$logLevel
    )
    $script:logLevel = $logLevel

    #log初期化
    $script:log = Join-Path $env:TEMP -ChildPath ('export_' + (Get-Date -Format 'yyyyMMddHHmmss') + '.log')
#    $script:log = [System.IO.Path]::GetTempFileName()
    LogInfo ""
    Start-Process powershell -ArgumentList "-NoExit", "-Command", "Get-Content -Path '$script:log' -Wait"

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
    Write-Output ( $time + ' ' + $level + ' [' + $func + '] ' + $message ) >> $script:log
}

