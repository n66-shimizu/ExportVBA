rem #chcp 65001
echo on
rem ####################################################
rem ##  tortoiseSVNのstart commit hockから連携され
rem ##  Access,ExcelのVBAソースをExportする 
rem ##  2026.06.20  PrimeBrains 清水
rem ####################################################
rem #  
rem #

rem #パラメーターを取得
set "REPOS=%~1"
set "SCRIPT_FOLDER=%~dp0"
set "SCRIPT_FOLDER=%SCRIPT_FOLDER%ps"

rem # office32/64ビット判定用のパス
rem # PC環境によって異なる。excelがインストール去れているパスを記述
set OFFICE32_PATH="C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE"
set OFFICE64_PATH="C:\Program Files\Microsoft Office\Office16\EXCEL.EXE"
rem # powershell32/64ビットのパス
rem # PC環境によって異なる。powershellがインストール去れているパスを記述
set PS32_PATH="C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
set PS64_PATH="C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

rem # officeが32bitか64bitか判定
if exist %OFFICE64_PATH% (
    echo 64ビット版PowerShellを起動します
    set PS=%PS64_PATH%
) else (
    echo 32ビット版PowerShellを起動します
    set PS=%PS32_PATH%
)

%PS% -ExecutionPolicy Bypass -Command "%SCRIPT_FOLDER%\ExportVBAForSVN.ps1 -repos '%REPOS%' -scriptFolder '%SCRIPT_FOLDER%'"
