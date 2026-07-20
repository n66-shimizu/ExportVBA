#
# Excel VBA Export
# ---------------------------------------------------
# ExportVBA.ps1 または、ExportVBAforSVN.ps1から実行
# 2026.06.20  PrimeBrains 清水
#

###############################
## 初期処理
###############################

function InitExportExcel {
    param (
        [string]$targetFile,
        [string]$exportFolder = ""
    )

    # targetFileの保存
    $script:targetFile = $targetFile

    # exportFolder(親フォルダ)取得
    $script:exportFolder = GetExportFolder -targetFile $script:targetFile -exportFolder $exportFolder
    LogInfo ("exportFolder=" + $script:exportFolder)
    InitFolder -folderPath $script:exportFolder

    # 拡張子判定
    $extension = [system.IO.Path]::GetExtension($script:targetFile).toLower()
    if ($extension -ne ".xlsm" -and $extension -ne ".xls") {
        LogError "$targetFile はExcel(macro有効)ファイルではありません。"
        Exit
    }

    # Excelアプリケーションを作成
    $script:excelApp = New-Object -ComObject Excel.Application
    $script:excelApp.Visible = $false

    # targetFileを開く
    $script:workbook = $script:excelApp.workbooks.Open($script:targetFile)
    $script:vbaProject = $script:workbook.VBProject

    if ($null -eq $script:vbaProject) {
        LogError @"
VBProject にアクセスできません。
Excel の設定で
「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」
が無効になっている可能性があります。
"@
        $script:workbook.Close($false)
        $script:excelApp.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:excelApp)
        Exit
    }
}

###############################
## 終了処理
###############################

function EndExportExcel {

    # workbookを閉じる
    $script:workbook.Close($false)

    # Excelアプリケーションを閉じる
    $script:excelApp.Quit()

    # Excel COMオブジェクトの解放
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:vbaProject)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:excelApp)

    # ガベージコレクションを実施してメモリを解放
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}


###############################
## エクスポート先フォルダ作成
###############################

function CreateExportExcelPath {

    # exportパス取得
    $script:exportPath = GetExportPath -targetFile $script:targetFile -exportFolder $script:exportFolder
    LogInfo ("script:exportPath=" + $script:exportPath)
    InitFolder -folderPath $script:exportPath

    # moduleフォルダの作成
    $script:modulePath = '050_MOD'
    $script:exportModulePath = Join-Path -Path $script:exportPath -ChildPath $script:modulePath
    InitFolderWithClear -folderPath $script:exportModulePath

    # classフォルダの作成
    $script:classPath = '010_CLS'
    $script:exportClassPath = Join-Path -Path $script:exportPath -ChildPath  $script:classPath
    InitFolderWithClear -folderPath $script:exportClassPath

    # formフォルダの作成
    $script:formPath = '020_FRM'
    $script:exportFormPath = Join-Path -Path $script:exportPath -ChildPath $script:formPath
    InitFolderWithClear -folderPath $script:exportFormPath

    # controlフォルダの作成
    $script:controlPath = '030_CTL'
    $script:exportControlPath = Join-Path -Path $script:exportPath -ChildPath  $script:controlPath
    InitFolderWithClear -folderPath $script:exportControlPath

}


###############################
## エクスポート処理
###############################

# コンポーネントのエクスポート
function DoExportComponent {

    foreach ($vbComponent in $script:vbaProject.VBComponents) {
         if ($vbComponent.Type -eq 1) {
            $compName = "$($vbComponent.Name).bas"
            $exportModulePath = Join-Path $script:exportModulePath -ChildPath "$compName"
            $vbComponent.Export($exportModulePath)
            LogInfo "Exported: $($script:modulePath)\$compName"
         } elseif ($vbComponent.Type -eq 100) {
            $compName = "$($vbComponent.Name).cls"
            $exportClassPath = Join-Path $script:exportClassPath -ChildPath "$compName"
            $vbComponent.Export($exportClassPath)
            LogInfo "Exported: $($script:classPath)\$compName"
         } elseif ($vbComponent.Type -eq 3) {
            $compName = "$($vbComponent.Name).frm"
            $exportFormPath = Join-Path $script:exportFormPath -ChildPath "$compName"
            $vbComponent.Export($exportFormPath)
            LogInfo "Exported: $($script:formPath)\$compName"
         } else {
            LogError "Others Name: $($vbComponent.Name), Type: $($vbComponent.Type)"
         }
    }
}

# コントロール情報の出力
function DoExportControl {

    foreach($sheet in $script:workbook.Sheets) {
        $compName = "$($sheet.Name).txt"
        $exportControlPath = Join-Path $script:exportControlPath -ChildPath "$compName"
        if (Test-Path $exportControlPath) {
            Clear-Content -Path $exportControlPath
        }
        # フォームコントロール
        $i = 0
        foreach ($shape in $sheet.Shapes) {
            if ($shape.Type -eq 8) { # 8 = Button (FormControl)
                if ($i -eq 0) {
                    Add-Content -Path $exportControlPath -Value 'Button: '
                    Add-Content -Path $exportControlPath -Value 'Number,Name,Caption,Left,Top,Width,Height,Action'
                }
                $i++;
                $name = $shape.Name
                $caption = $shape.TextFrame.Characters().Text 
                $caption = $caption.Replace("`r`n", "\r\n")
                $left = $shape.Left
                $top = $shape.Top
                $width = $shape.Width
                $height = $shape.Height
                $action = $shape.OnAction
                #Add-Content -Path $exportControlPath -Value ( "" + $shape.Type )
                Add-Content -Path $exportControlPath -Value ( "" + $i + ',"' + $name + '","' + $caption + '",' + $left + ',' + $top + ',' + $width + ',' + $height + ',"' + $action + '"' )
            }
        }
        # 空行出力
        if ($i -gt 0) {
            Add-Content -Path $exportControlPath -Value ''
        }
        # ActiveX コントロール
        $j = 0
        foreach ($ole in $sheet.OLEObjects()) {
            if ($ole.OLEType -eq 2) {
                try {
                    $control = $ole.Object
                    #if ($control.GetType().Name -eq "CommandButton") {
                        if ($j -eq 0) {
                            Add-Content -Path $exportControlPath -Value 'ActiveX Button: '
                            Add-Content -Path $exportControlPath -Value 'Number,Name,Caption,Left,Top,Width,Height,Action'
                        }
                        $j++;                        
                        $name = $control.Name
                        $caption = $control.Caption.Replace("`r`n", "\r\n")
                        $left = $ole.Left
                        $top = $ole.Top
                        $width = $ole.Width
                        $height = $ole.Height
                        $action = '?'
                        
                        #イベントハンドラが置かれているモジュールを仮定
                        $moduleName = $sheet.CodeName
                        $eventName = "$($control.Name)_Click"
                        foreach ($vbComponent in $script:vbaProject.VBComponents) {
                            $codeModule = $vbComponent.CodeModule
                            $lines = $codeModule.lines(1, $codeMOdule.CountOfLines)
                            
                            if ($lines -match $eventName) {
                                $action = $eventName
                                break
                            }
                        }
                        Add-Content -Path $exportControlPath -Value ( "" + $j + ',"' + $name + '","' + $caption + '",' + $left + ',' + $top + ',' + $width + ',' + $height + ',"' + $action + '"' )
                    #}
                } catch {
                
                }
            }
        }
        if (($i -gt 0) -or ($j -gt 0)) {
            LogInfo "Exported: $($script:exportControl)\$compName"
        }
        
    }
}

# Excel全体のエクスポート
function DoExportExcel {

    DoExportComponent		
    DoExportControl

    LogInfo "$($script:targetFile) All VBA Componets have been Exported successfully."
}

###############################
## 主処理
###############################

function exportExcelVBA {
    param (
        [string]$targetFile,
        [string]$exportFolder
    )

    # exportの準備
    InitExportExcel -targetFile $targetFile -exportFolder $exportFolder

    # exportフォルダを準備
    CreateExportExcelPath

    # exportを実施
    doExportExcel

    # exportの後始末
    EndExportExcel
}

