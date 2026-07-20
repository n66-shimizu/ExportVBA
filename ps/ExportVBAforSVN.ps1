#####################################################
# ExportVBAForSVN.batと連携して
# Access,ExcelのVBAソースをExportする 
# 2026.06.20  PrimeBrains 清水
##################################################### 
#
#

###############################
## パラメーター
###############################

param (
    [string]$repos,
    [string]$scriptFolder
)


##############################################
## 拡張子を分別して
## Access用またはExcel用Exportを実行
###############################################

function ExportByExtension {
    param (
        [string]$targetFile
    )

    $targetLeaf = Split-Path -Path $targetFile -Leaf
    $extension =  [system.IO.Path]::GetExtension("$targetFile").toLower()

    if (($extension -eq '.xlsm') -or ($extension -eq '.xls')) {
        LogInfo "Export: $targetLeaf (Excel)"
        exportExcelVBA -targetFile  $targetFile
    } elseif (($extension -eq '.accdb') -or ($extension -eq '.mdb')) {
        LogInfo "Export: $targetLeaf (Access)"
        exportAccessVBA -targetFile  $targetFile
    }
}

##############################################
## 実行確認
###############################################
function confirmExport
{
    Add-Type -AssemblyName System.Windows.Forms
    
    #メッセージボックスを表示
    $result = [System.Windows.Forms.MessageBox]::Show(
        "VBA Exportを実行しますか",
        "確認",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    #回答判定
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        return $true
    } 
    Exit
}

##############################################
## 主処理
###############################################

function exportVBAforSYN {
    param (
        [string]$repos,
        [string]$scriptFolder
    )


    #実行確認
    confirmExport
    
    #scriptパス
    $script:scriptFolder = $scriptFolder

    $commonPath = Join-Path $script:scriptFolder -ChildPath 'Common.ps1'
    . "$commonPath"

    $commonLogPath = Join-Path $script:scriptFolder -ChildPath 'CommonWindowLog.ps1'
    . "$commonLogPath"
    #log初期化
    InitWindowLog 9
    LogInfo "Include: $commonPath" 
    LogInfo "Include: $commonLogPath" 
    $scriptPath = Join-Path $script:scriptFolder -ChildPath 'ExportExcelVBA.ps1'
    LogInfo "Include: $scriptPath" 
    . "$scriptPath"

    $scriptPath = Join-Path $script:scriptFolder -ChildPath 'ExportAccessVBA.ps1'
    LogInfo "Include: $scriptPath" 
    . "$scriptPath"

    #repos読み取り ファイル渡しのパラメータを読み込む
    $Script:reposPath = Get-Content -Path $repos -Encoding UTF8
    LogInfo "Repos: $Script:reposPath"
     
    # svn statusを実行
    $svnStatus = svn status
    #LogDebug "$svnStatus" 

    #１行ずつ処理
    foreach ($line in $svnStatus) {
        LogInfo $line 
        $status     = $line.Substring(0,1)
        $targetFile = $line.Substring(8).trim()
        $targetFile = Join-Path $Script:reposPath -ChildPath $targetFile
    
        if (($status -eq 'A') -or ($status -eq 'M') -or ($status -eq '?')) {
            ExportByExtension -targetFile $targetFile
        } elseif ($status -eq '!') {   # 消された場合、exportも削除する
            $exportFolder = GetExportFolder -targetFile $targetFile
            $exportPath = GetExportPath -targetFile $targetFile -exportFolder $exportFolder
            if (Test-Path $exportPath) {
                Remove-Item -Path $exportPath -Recurse -Force
                LogInfo "Remove: $exportPath"
            }
        }
    }
    
    #終了
    EndLog

}
    
###############################
## スクリプト本体
###############################

exportVBAforSYN -repos $repos -scriptFolder $scriptFolder



