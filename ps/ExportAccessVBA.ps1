#
# Access VBA Export
# ---------------------------------------------------
# ExportVBA.ps1 または、ExportVBAforSVN.ps1から実行
# 2026.06.20  PrimeBrains 清水
#

###############################
## Keyboardイベント
###############################

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class Keyboard {
    [DllImport("user32.dll")]
    public static extern void keybd_event(
        byte bVk,
        byte bScan,
        int dwFlgas,
        int dwExtraInfo);

    public const int KEYEVENTF_KEYUP = 0x0002;
}
"@

$VK_SHIFT  = 0x10
$VK_LSHIFT = 0xA0
$VK_RSHIFT = 0xA1

###############################
## 初期処理
###############################

function InitExportAccess {
    param (
        [string]$targetFile,
        [string]$exportFolder = ""
    )

    LogInfo "初期処理"

    # targetFileの保存
    $script:targetFile = $targetFile

    # exportFolder(親フォルダ)取得
    $script:exportFolder = GetExportFolder -targetFile $script:targetFile -exportFolder $exportFolder
    LogInfo ("exportFolder=" + $script:exportFolder)
    InitFolder -folderPath $script:exportFolder

    # 拡張子判定
    $extension = [system.IO.Path]::GetExtension($script:targetFile).toLower()
    if ($extension -ne ".accdb" -and $extension -ne ".mdb") {
        LogError "$targetFile はAccessファイルではありません。"
        Exit
    }

    # Accessアプリケーションを作成
    $script:accessApp = New-Object -ComObject Access.Application
    $script:accessApp.Visible = $false

    try {
        # Shift Key On エミュレート
        [Keyboard]::keybd_event($VK_SHIFT,0,0,0)
        Start-Sleep -Milliseconds 300

        # Accessデータベースを開く
        $script:accessApp.OpenCurrentDatabase($script:targetFile, $false)
        Start-Sleep -Seconds 2
    }
    finally {
        # Shift Key Up エミュレート
        [Keyboard]::keybd_event($VK_SHIFT,0,[Keyboard]::KEYEVENTF_KEYUP,0)
        [Keyboard]::keybd_event($VK_LSHIFT,0,[Keyboard]::KEYEVENTF_KEYUP,0)
        [Keyboard]::keybd_event($VK_RSHIFT,0,[Keyboard]::KEYEVENTF_KEYUP,0)
    }

    # DAOによるDBの取得
    $dao = New-Object -ComObject "DAO.DBEngine.120" # Access 2010以降は"120"
    $script:db = $dao.OpenDatabase($script:targetFile)

    # VBAプロジェクトへのアクセスを取得
    $script:vbaProject = $script:accessApp.VBE.VBprojects.Item(1)

}

###############################
## 終了処理
###############################

function EndExportAccess {
    LogInfo "終了処理"

    # dbを閉じる
    $script:db.Close()

    # Accessアプリケーションを終了する
    $script:accessApp.CloseCurrentDatabase()
    $script:accessApp.Quit()

    # Access COMオブジェクトの解放
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:accessApp)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:db)

    $script:vbaProject = $null
    $script:db = $null
    $script:accessApp = $null

    # ガベージコレクションを実施してメモリを解放
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # Shift Keyup が効かない場合に強制リセット
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait("+A")

}


###############################
## エクスポート先フォルダ作成
###############################

function CreateExportAccessPath {

    # exportパス取得
    $script:exportPath = GetExportPath -targetFile $script:targetFile -exportFolder $script:exportFolder
    LogInfo ("script:exportPath=" + $script:exportPath)
    InitFolder -folderPath $script:exportPath

    # tableフォルダの作成
    $script:tablePath = "010_TBL"
    $script:exportTablePath = Join-Path -Path $script:exportPath -ChildPath $script:tablePath
    InitFolderWithClear -folderPath $script:exportTablePath

    # link tableフォルダの作成
    $script:linkTablePath = "011_LINK"
    $script:exportLinkTablePath = Join-Path -Path $script:exportPath -ChildPath $script:linkTablePath
    InitFolderWithClear -folderPath $script:exportLinkTablePath

    # sysem tableフォルダの作成
    $script:sysTablePath = "012_SYS"
    $script:exportSysTablePath = Join-Path -Path $script:exportPath -ChildPath $script:sysTablePath
    InitFolderWithClear -folderPath $script:exportSysTablePath

    # formフォルダの作成
    $script:formPath = "020_FRM"
    $script:exportFormPath = Join-Path -Path $script:exportPath -ChildPath $script:formPath
    InitFolderWithClear -folderPath $script:exportFormPath

    # reportフォルダの作成
    $script:reportPath = "030_REP"
    $script:exportReportPath = Join-Path -Path $script:exportPath -ChildPath $script:reportPath
    InitFolderWithClear -folderPath $script:exportReportPath

    # macroフォルダの作成
    $script:macroPath = "040_MCR"
    $script:exportMacroPath = Join-Path -Path $script:exportPath -ChildPath $script:macroPath
    InitFolderWithClear -folderPath $script:exportMacroPath

    # moduleフォルダの作成
    $script:modulePath = "050_MOD"
    $script:exportModulePath = Join-Path -Path $script:exportPath -ChildPath  $script:modulePath
    InitFolderWithClear -folderPath $script:exportModulePath

    # queryフォルダの作成
    $script:queryPath = "060_QRY"
    $script:exportQueryPath = Join-Path -Path $script:exportPath -ChildPath $script:queryPath
    InitFolderWithClear -folderPath $script:exportQueryPath

    # sqlフォルダの作成
    $script:sqlPath = "061_SQL"
    $script:exportSqlPath = Join-Path -Path $script:exportPath -ChildPath $script:sqlPath
    InitFolderWithClear -folderPath $script:exportSqlPath

}

###############################
## エクスポート処理
###############################

# テーブルのfield行を整形する
 function GetFieldInfo {
     param (
        [int]$number,
        [object]$field
    )

    $fieldName = $field.Name
    $fieldSize = $field.Size
  
    switch($field.Type) {
        1   { $fieldJType = "ブール型"}
        2   { $fieldJType = "バイト型"}
        3   { $fieldJType = "整数型"}
        4   { $fieldJType = "長整数型"}
        5   { $fieldJType = "通貨型"}
        6   { $fieldJType = "単精度浮動小数点数型"}
        7   { $fieldJType = "倍精度浮動小数点数型"}
        8   { $fieldJType = "日付・時刻型"}
        9   { $fieldJType = "バイト型"}
        10  { $fieldJType = "テキスト型"}
        11  { $fieldJType = "ロングバイナリー型"}
        12  { $fieldJType = "メモ型"}
        15  { $fieldJType = "GUID型"}
        16  { $fieldJType = "BigInt型"}
        17  { $fieldJType = "可変長バイナリー型"}
        18  { $fieldJType = "Char型"}
        19  { $fieldJType = "Numeric型"}
        20  { $fieldJType = "10進型"}
        21  { $fieldJType = "浮動小数点数型"}
        22  { $fieldJType = "時刻型"}
        23  { $fieldJType = "タイムスタンプ型"}
        101 { $fieldJType = "添付ファイル"}
        102 { $fieldJType = "複合列（Byte）"}
        103 { $fieldJType = "複合列（Integer）"}
        104 { $fieldJType = "複合列（Long）"}
        105 { $fieldJType = "複合列（Single）"}
        106 { $fieldJType = "複合列（Double）"}
        107 { $fieldJType = "複合列（GUID）"}
        108 { $fieldJType = "複合列（Decimal）"}
        109 { $fieldJType = "複合列（Text）"}
        110 { $fieldJType = "複合列（Memo）"}
        default { $fieldJType = "その他"} 
    }

    switch($field.Type) {
        1   { $fieldType = "dbBoolean"}
        2   { $fieldType = "dbByte"}
        3   { $fieldType = "dbInteger"}
        4   { $fieldType = "dbLong"}
        5   { $fieldType = "dbCurrency"}
        6   { $fieldType = "dbSingle"}
        7   { $fieldType = "dbDouble"}
        8   { $fieldType = "dbDate"}
        9   { $fieldType = "dbBinary"}
        10  { $fieldType = "dbText"}
        11  { $fieldType = "dbLongBinary"}
        12  { $fieldType = "dbMemo"}
        15  { $fieldType = "dbGUID"}
        16  { $fieldType = "dbBigInt"}
        17  { $fieldType = "dbVarBinary"}
        18  { $fieldType = "dbChar"}
        19  { $fieldType = "dbNumeric"}
        20  { $fieldType = "dbDecimal"}
        21  { $fieldType = "dbFloat"}
        22  { $fieldType = "dbTime"}
        23  { $fieldType = "dbTimeStamp"}
        101 { $fieldType = "dbAttachment"}
        102 { $fieldType = "dbComplexByte"}
        103 { $fieldType = "dbComplexInteger"}
        104 { $fieldType = "dbComplexLong"}
        105 { $fieldType = "dbComplexSingle"}
        106 { $fieldType = "dbComplexDouble"}
        107 { $fieldType = "dbComplexGUID"}
        108 { $fieldType = "dbComplexDecimal"}
        109 { $fieldType = "dbComplexText"}
        110 { $fieldType = "dbComplexMemo"}
        default { $fieldType = "その他"} 
    }
    
    #$field | Get-Member
    $fieldRequired = $field.Required
 
    return  '' + $number + ',"' + $fieldName + '",' +  $fieldSize + ',' + $fieldType + ',' + $fieldRequired
}

# テーブルのIndex行を整形する
 function GetIndexInfo {
     param (
        [int]$number,
        [object]$index
    )

    $indexName = $index.Name
    $fields = ($index.Fields | ForEach-Object { $_.Name }) -join ", " 
    if ($index.Primary) {
        $primaryKey = "True"
    } else {
        $primaryKey = "False"
    }
    $required = $index.Required
    $unique =  $index.Unique
 
    #$index | Get-Member
    return  '' + $number + ',"' + $indexName + '","' +  $fields + '",' + $primaryKey + ',' + $required + ',' + $unique
}

# １テーブル情報のエクスポート
function DoExportTableInfo {
    param (
        [string]$exportPath,
        [object]$table
    )
    $values = @()
    $values += ('Name   : ' + $table.Name)
    $values += ('Connect: ' + $table.Connect)
    $values += ('LastUpdated : ' + $table.LastUpdated) 
    
    # fields
    if ( $table.Fields.Count -gt 0) {
         $values += ''
         $values += 'Fields  : '
         $values += 'Number,Name,Size,Type,NotNull'
         for($i = 0; $i -lt $table.Fields.Count; $i++) {
             $values += ( GetFieldInfo -Number ($i+1) -Field $table.Fields($i) )
         }
    }
    
    # indexes
    if ( $table.Indexes.Count -gt 0) {
         $values += ''
         $values += 'Indexes : '
         $values += 'Number,Name,Fields,PrimaryKey,Required,Unique'
         for($i = 0; $i -lt $table.Indexes.Count; $i++) {
             $values += ( GetIndexInfo -Number ($i+1) -Index $table.Indexes($i) )
         }
    }
    Set-Content -Path $exportPath -Value $values
}

# テーブルのエクスポート
function DoExportTable {
    $tables = $script:db.TableDefs | Where-Object { $_.Name -notlike "MSys*" -and $_.Attributes -eq 0 } 
    foreach ( $table in $tables ) {
        if (-not ($table.Name.startsWith("~"))) {
            $exportTablePath = Join-Path $script:exportTablePath -ChildPath "$($table.Name).txt"
            DoExportTableInfo -exportPath $exportTablePath -table $table
            LogInfo "Exported: $($script:tablePath)\$($table.Name).txt"
        }
    }
}


# リンクテーブルのエクスポート
function DoExportLinkTable {
    $linkTables = $script:db.TableDefs | Where-Object { $_.Name -notlike "MSys*" -and ( $_.Attributes -eq 1073741824 -or  $_.Attributes -eq 537001984 ) } 
    foreach ( $linkTable in $linkTables ) {
        if (-not ($linkTable.Name.startsWith("~"))) {
            $exportLinkTablePath = Join-Path $script:exportLinkTablePath -ChildPath "$($linkTable.Name).txt"
            DoExportTableInfo -exportPath $exportLinkTablePath -table $linkTable
            LogInfo "Exported: $($script:linkTablePath)\$($linkTable.Name).txt"
        }
    }
}

# システムテーブルのエクスポート
function DoExportSysTable {
    $sysTables = $script:db.TableDefs | Where-Object { $_.Name -like "MSys*" } 
    foreach ( $sysTable in $sysTables ) {
        if (-not ($sysTable.Name.startsWith("~"))) {
            $exportSysTablePath = Join-Path $script:exportSysTablePath -ChildPath "$($sysTable.Name).txt"
            DoExportTableInfo -exportPath $exportSysTablePath -table $sysTable
            LogInfo "Exported: $($script:sysTablePath)\$($sysTable.Name).txt"
        }
    }
}

# VBAモジュールのエクスポート
function DoExportModule {

    # 標準モジュール
    foreach ($vbComponent in $script:vbaProject.VBComponents) {
        if ($vbComponent.Type -eq [Microsoft.Vbe.Interop.vbext_ComponentType]::vbext_ct_StdModule) { # Standard module
            $exportModulePath = Join-Path $script:exportModulePath -ChildPath "$($vbComponent.Name).bas" 
            $vbComponent.Export($exportModulePath)
            LogInfo "Exported: $($script:modulePath)\$($vbComponent.Name).bas"
        } 
    }

    # clsモジュール
    foreach ($vbComponent in $script:vbaProject.VBComponents) {
        if ($vbComponent.Type -eq [Microsoft.Vbe.Interop.vbext_ComponentType]::vbext_ct_ClassModule) { # Standard module
            $exportModulePath = Join-Path $script:exportModulePath -ChildPath "$($vbComponent.Name).cls" 
            $vbComponent.Export($exportModulePath)
            LogInfo "Exported: $($script:modulePath)\$($vbComponent.Name).cls"
        } 
    }

    # フォーム、レポートモジュール
    # CurrentプロジェクトからのExportにVBAコード部分も含まれるので、重複をさけるためvbaからのエクスポートは抑制
    ##foreach ($vbComponent in $script:vbaProject.VBComponents) {
    ##    if ($vbComponent.Type -eq [Microsoft.Vbe.Interop.vbext_ComponentType]::vbext_ct_Document) { # Standard module
    ##        $exportModulePath = Join-Path $script:exportModulePath -ChildPath "$($vbComponent.Name).cls" 
    ##        $vbComponent.Export($exportModulePath)
    ##        LogInfo "Exported: $($script:modulePath)\$($vbComponent.Name).cls"
    ##    } 
    ##}

}

# SQLのエクスポート
function DoExportSql {
    $queryDefs = $script:db.QueryDefs
    foreach($queryDef in $queryDefs) {
       if (-not ($queryDef.Name.startsWith("~"))) {  # フォーム、レポートのパーツのSQLを出力しない
           $exportSqlPath = Join-Path $script:exportSqlPath -ChildPath "$($queryDef.Name).sql"
           Set-Content -Path $exportSqlPath -value $queryDef.SQL
           LogInfo "Exported: $($script:sqlPath)\$($queryDef.Name).txt"
       }
    }   
}

# マクロのエクスポート
function DoExportMacro {
    $macros = $script:accessApp.CurrentProject.AllMacros
    foreach ($macro  in $macros ) {
        $exportMacroPath = Join-Path $script:exportMacroPath -ChildPath "$($macro.Name).txt"
        $script:accessApp.SaveAsText(4, $macro.Name, $exportMacroPath)
        LogInfo "Exported: $($script:macroPath)\$($macro.Name).txt"
    }
 }

# クエリのエクスポート
function DoExportQuery {
    $queries = $script:accessApp.CurrentData.AllQueries
    foreach ( $query in $queries ) {
        $exportQueryPath = Join-Path $script:exportQueryPath -ChildPath "$($query.Name).txt"
        $script:accessApp.SaveAsText(1, $query.Name, $exportQueryPath)
        LogInfo "Exported: $($script:queryPath)\$($query.Name).txt"
    }
 }

# フォームのエクスポート
function DoExportForm {
    $forms = $script:accessApp.CurrentProject.AllForms
    foreach ( $form in $forms ) {
        $exportFormPath = Join-Path $script:exportFormPath -ChildPath "$($form.Name).frm"    
        $accessApp.SaveAsText(2, $form.Name, $exportFormPath)
        LogInfo "Exported: $($script:formPath)\$($form.Name).txt"
    }
}

# レポートのエクスポート
function DoExportReport {
    $reports = $script:accessApp.CurrentProject.AllReports
    foreach ( $report in $reports ) {
        $exportReportPath = Join-Path $script:exportReportPath -ChildPath "$($report.Name).cls"    
        $script:accessApp.SaveAsText(3, $report.Name, $exportReportPath)
        LogInfo "Exported: $($script:reportPath)\$($report.Name).txt"
    }
}

# Access全体のエクスポート
function DoExportAccess {

    DoExportTable
    DoExportLinkTable
    DoExportSysTable
    DoExportForm
    DoExportReport
    DoExportMacro
    DoExportModule
    DoExportQuery
    DoExportSql

    LogInfo "$($script:targetFile) All VBA Componets have been Exported successfully."
}

###############################
## 主処理
###############################

function exportAccessVBA {
    param (
        [string]$targetFile,
        [string]$exportFolder
    )

    # exportの準備
    InitExportAccess -targetFile $targetFile -exportFolder $exportFolder

    # exportフォルダを準備
    CreateExportAccessPath

    # exportを実施
    doExportAccess

    # exportの後始末
    EndExportAccess
}



