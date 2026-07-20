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
    [string]$targetFile,
    [string]$scriptFolder,
    [string]$exportFolder = ""
)


##############################################
## 主処理
###############################################

function ExportVBA {
    param (
        [string]$targetFile,
        [string]$scriptFolder,
        [string]$exportFolder = ""
    )

    #scriptパス
    $script:scriptFolder = $scriptFolder

    $commonPath = Join-Path $script:scriptFolder -ChildPath 'Common.ps1'
    . "$commonPath"
    #log初期化
    InitLog 9
    LogInfo "Include: $commonPath" 

    $extension =  [system.IO.Path]::GetExtension("$targetFile").toLower()
    if (($extension -eq '.xlsm') -or ($extension -eq '.xls')) {
       $scriptPath = Join-Path $script:scriptFolder -ChildPath 'ExportExcelVBA.ps1'
       LogInfo "Include: $scriptPath"         
       . "$scriptPath"
       exportExcelVBA -targetFile $targetFile -exportFolder $exportFolder
    } elseif (($extension -eq '.accdb') -or ($extension -eq '.mdb')) {
       $scriptPath = Join-Path $script:scriptFolder -ChildPath 'ExportAccessVBA.ps1'
       LogInfo "Include: $scriptPath" 
       . "$scriptPath" 
       exportAccessVBA -targetFile $targetFile -exportFolder $exportFolder
    }
    
    EndLog
}

###############################
## スクリプト本体
###############################

ExportVBA -targetFile $targetFile -scriptFolder $scriptFolder -exportFolder $exportFolder

