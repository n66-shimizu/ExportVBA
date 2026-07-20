# Access/Excel VBAエクスポートツール 説明

## 1. 概要

このツールは、**Microsoft AccessファイルおよびExcel(マクロ有効)ファイルの中身（VBAソースコード、テーブル定義、フォーム、クエリ等）をテキストファイルとして外部にエクスポート**するためのバッチファイル＋PowerShellスクリプトの組み合わせです。

主な目的は、Access/Excelのバイナリファイル（.accdb, .mdb, .xlsm, .xls）に含まれるVBAコードやオブジェクト定義を**テキスト化してバージョン管理（Git/SVNなど）で差分管理できるようにする**ことです。

大きく分けて2つの使い方が用意されています。

| 用途 | エントリポイント | 想定シーン |
|---|---|---|
| 手動で1ファイルをエクスポート | `ExportVBA.bat` | エクスプローラーからファイルをドラッグ＆ドロップして単発実行 |
| SVNのコミット時に自動連携してエクスポート | `ExportVBAforSVN.bat` | TortoiseSVNの「start-commitフック」から呼び出され、変更のあった全ファイルを自動エクスポート |

---

## 2. ファイル構成と役割

```
ExportVBA.bat            … 手動実行用の入り口（バッチ）
ExportVBAforSVN.bat      … SVN連携用の入り口（バッチ）
ps/
 ├ ExportVBA.ps1          … 単一ファイルをエクスポートするメイン処理
 ├ ExportVBAforSVN.ps1    … SVNのstatusを見て変更ファイルを一括エクスポートするメイン処理
 ├ ExportExcelVBA.ps1     … Excelファイル向けのエクスポート処理本体
 ├ ExportAccessVBA.ps1    … Accessファイル向けのエクスポート処理本体
 ├ Common.ps1             … 共通関数（フォルダ操作、ログ出力 等）
 └ CommonWindowLog.ps1    … ログを別ウィンドウにリアルタイム表示するための拡張（SVN連携時に使用）
```

※ `.ps1`は `.bat` と同じ階層の `ps` フォルダに置く前提になっています。

---

## 3. 使用方法

### 3-1. ExportVBA.bat（手動実行用）

- 第1引数 `%1` に対象ファイル（.xlsm/.xls/.accdb/.mdbのパス）を受け取る
- PCにインストールされているOfficeが32bit版か64bit版かを、Excel実行ファイルの有無で判定
  - 64bit Officeがあれば64bit版PowerShellを、なければ32bit版PowerShellを使用
  - ※COMオブジェクト（Excel.Application等）を扱うため、**PowerShellのbit数とOfficeのbit数を一致させる必要がある**ための処理
- 判定したPowerShellで `ExportVBA.ps1` を `-ExecutionPolicy Bypass` で実行
- 最後に `pause` があるため、実行結果を確認してからウィンドウを閉じる

**使い方（想定）**
```
ExportVBA.bat "C:\path\to\対象ファイル.xlsm"
```
- または、対象ファイルを `ExportVBA.bat` のアイコンにドラッグ＆ドロップして実行する。
- 対象ファイルを右クリックしてコンテキストメニューを開き、”プログラムから開く..."から`ExportVBA.bat`を選択する。

### 3-2. ExportVBAforSVN.bat（SVN連携用）

- 第1引数 `%1` に**SVNのリポジトリのローカルパスを記載したファイル**を受け取る（TortoiseSVNのフックスクリプトから渡される想定）
- 同様にOffice/PowerShellのbit数判定を行う
- `ExportVBAforSVN.ps1` を実行する

**使い方（想定）**
TortoiseSVNの「設定 > フックスクリプト」で、**start-commitフック**として以下のように登録する。
```
ExportVBAforSVN.bat "%PATHFILE%"
```
（`%PATHFILE%` はTortoiseSVNがコミット対象パス一覧を書き込む一時ファイルへの参照）

これにより、**コミットしようとした瞬間に、変更のあったAccess/Excelファイルの中身が自動でテキスト化され、それも一緒にコミットできる**ようになります。

## 4. エクスポート先フォルダの構造

対象ファイルと同じ階層に `ExportedVBA` フォルダが作られ（未指定の場合）、その下に対象ファイルごとのサブフォルダが作られます。

例：`C:\project\業務システム.accdb` をエクスポートすると

```
C:\project\ExportedVBA\
  └ 業務システム_accdb\
      ├ 010_TBL\      … テーブル定義（.txt）
      ├ 011_LINK\     … リンクテーブル定義
      ├ 012_SYS\      … システムテーブル定義
      ├ 020_FRM\      … フォーム
      ├ 030_REP\      … レポート
      ├ 040_MCR\      … マクロ
      ├ 050_MOD\      … 標準/クラスモジュール
      ├ 060_QRY\      … クエリ
      └ 061_SQL\      … SQL文
```

Excelの場合は `010_CLS` `020_FRM` `030_CTL` `050_MOD` のみが作られます。

---

## 5. 事前準備・注意点

1. **Excel/Accessそれぞれで「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」設定を有効にする**必要があります。
   （Excel: ファイル > オプション > トラストセンター > マクロの設定）
   これが無効だとExcelエクスポート時にエラーになります。
2. `ExportVBA.bat` 冒頭の変数 `OFFICE32_PATH` / `OFFICE64_PATH` / `PS32_PATH` / `PS64_PATH` に、OfiiceおよびPowershellのバージョン、実行ファイルパスを記載する必要があります。**実行するPCの環境に合わせて書き換えて下さい**。
3. Access/Excelを裏で（非表示で）起動して自動操作するため、実行中は対象ファイルを他で開いていないこと。処理中にダイアログ（保存確認等）が出るとスクリプトが止まる可能性があります。
4. Accessのエクスポートでは、Shiftキーの疑似押下によりAutoExecマクロの実行を抑止していますが、環境によっては動作しない場合があります。
5. SVN連携版は**TortoiseSVNのstart-commitフック**として使う前提であり、単体で `.bat` を叩く場合は引数にコミット対象パス一覧が書かれたテキストファイルを渡す必要があります。
6. ログレベルはコード内で `9`（デバッグ最大）に固定されています。ログを絞りたい場合は `InitLog` / `InitWindowLog` の呼び出し部分の数値を変更してください（`LogInfo`は2以上、`LogDebug`は9以上で出力）。

---

## 6. 全体の処理フロー図

```
[手動実行]                          [SVNコミット時]
ExportVBA.bat                       ExportVBAforSVN.bat
   │ (対象ファイル1つ)                  │ (パス一覧ファイル)
   ▼                                    ▼
ExportVBA.ps1                       ExportVBAforSVN.ps1
   │                                    │ 確認ダイアログ表示
   │ 拡張子判定                          │ svn status 実行
   │                                    │ 変更行ごとに拡張子判定
   ▼                                    ▼
 ┌─────────────┬─────────────┐
 │ ExportExcelVBA.ps1 │ ExportAccessVBA.ps1 │
 │ (.xlsm/.xls)       │ (.accdb/.mdb)        │
 └─────────────┴─────────────┘
        │
        ▼
  ExportedVBA配下にテキスト出力
  （Common.ps1のログ・フォルダ関数を共通利用）
```

## 7 各Powershellの機能詳細
### 7-1. ExportVBA.ps1（単一ファイル用メイン処理）

`ExportVBA.bat` から呼ばれる本体です。

1. `Common.ps1` を読み込み、ログレベル9（デバッグ）でログ初期化
2. 対象ファイルの拡張子を判定
   - `.xlsm` / `.xls` → `ExportExcelVBA.ps1` を読み込み `exportExcelVBA` を実行
   - `.accdb` / `.mdb` → `ExportAccessVBA.ps1` を読み込み `exportAccessVBA` を実行

### 7-2. ExportVBAforSVN.ps1（SVN連携用メイン処理）

`ExportVBAforSVN.bat` から呼ばれる本体です。

1. `confirmExport`：**「VBA Exportを実行しますか」という確認ダイアログ**を表示し、「いいえ」ならその場で処理を中断（＝コミットも中断される想定）
2. `Common.ps1` と `CommonWindowLog.ps1` を読み込み、別ウィンドウ形式でログ初期化
3. `ExportExcelVBA.ps1` / `ExportAccessVBA.ps1` の両方を読み込む
4. 引数で渡されたファイル（TortoiseSVNが用意したパス一覧ファイル）から**リポジトリのローカルパス**を読み取る
5. `svn status` を実行し、変更のあった全ファイルを1行ずつ処理
   - ステータスが `A`（追加）・`M`（変更）・`?`（未管理）→ 拡張子に応じてExcel/Accessのエクスポートを実行
   - ステータスが `!`（欠落＝削除された）→ 対応するエクスポート済みフォルダも削除（ソースの削除に追従）

### 7-3. ExportExcelVBA.ps1（Excel用エクスポート処理）

Excelファイル（.xlsm/.xls）からVBAとコントロール情報を抽出します。

**処理の流れ（`exportExcelVBA`）**
1. `InitExportExcel`：拡張子チェック → Excelを非表示で起動 → 対象ファイルを開く → `VBProject` を取得
   - VBProjectが取得できない場合、「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」設定が無効である旨のエラーを出して終了
2. `CreateExportExcelPath`：エクスポート先に以下のサブフォルダを作成（既存があればクリア）
   - `010_CLS`（クラスモジュール）
   - `020_FRM`（フォーム）
   - `030_CTL`（コントロール情報）
   - `050_MOD`（標準モジュール）
3. `DoExportComponent`：VBAコンポーネントをタイプ別に `.bas`（標準モジュール）/ `.cls`（クラス）/ `.frm`（フォーム）としてエクスポート
4. `DoExportControl`：各シート上の**フォームコントロールのボタン**および**ActiveXコントロールのボタン**の情報（名前・キャプション・位置・サイズ・実行マクロ名）をシートごとに `.txt` として出力
5. `EndExportExcel`：ブックを閉じてExcelを終了、COMオブジェクトを解放しガベージコレクションを実行

### 7-4. ExportAccessVBA.ps1（Access用エクスポート処理）

Accessファイル（.accdb/.mdb）から、テーブル定義・フォーム・レポート・マクロ・クエリ・SQL・VBAモジュールを幅広くエクスポートします。

**特徴的な前処理**
- Accessを開く際、**Shiftキー押下をキーボードAPI（`keybd_event`）でエミュレートしながら開く**処理があります。これはAccessの「自動実行（AutoExec）マクロ」の実行をShiftキーで抑止するための一般的なテクニックです。
- `DAO.DBEngine.120` を使ってDAO経由でもデータベースを開き、テーブル定義情報の取得に利用しています。

**処理の流れ（`exportAccessVBA`）**
1. `InitExportAccess`：拡張子チェック → Shiftキーを押しながらAccessを非表示で起動しDBを開く → DAOでもDBを開く → VBAプロジェクトを取得
2. `CreateExportAccessPath`：エクスポート先に以下のサブフォルダを作成
   - `010_TBL`（通常テーブル）
   - `011_LINK`（リンクテーブル）
   - `012_SYS`（システムテーブル `MSys*`）
   - `020_FRM`（フォーム）
   - `030_REP`（レポート）
   - `040_MCR`（マクロ）
   - `050_MOD`（VBAモジュール）
   - `060_QRY`（クエリ）
   - `061_SQL`（クエリのSQL文）
3. 各`DoExport〜`関数でエクスポートを実行
   - `DoExportTable` / `DoExportLinkTable` / `DoExportSysTable`：テーブルのフィールド定義・インデックス定義を日本語型名付きで `.txt` に出力
   - `DoExportForm` / `DoExportReport`：`SaveAsText` でテキスト形式にエクスポート
   - `DoExportMacro`：`SaveAsText` でマクロをエクスポート
   - `DoExportModule`：標準モジュール（`.bas`）とクラスモジュール（`.cls`）をエクスポート（フォーム/レポート内蔵のVBAコードは、フォーム/レポート自体のエクスポートに含まれるため二重出力を回避し除外）
   - `DoExportQuery`：クエリを `SaveAsText` でエクスポート
   - `DoExportSql`：各クエリのSQL文を `.sql` として出力（フォーム/レポートの内部クエリ `~`始まりは除外）
4. `EndExportAccess`：DBを閉じてAccessを終了、COMオブジェクトを解放。念のため `SendKeys` で Shift キー状態を強制リセット

### 7-5. Common.ps1（共通関数）

全スクリプトから読み込まれる共通関数集です。

| 関数名 | 機能 |
|---|---|
| `InitFolder` | 指定フォルダが存在しなければ作成する |
| `InitFolderWithClear` | 指定フォルダが存在すれば中身を全削除、なければ新規作成する（エクスポート先を毎回まっさらにするため） |
| `ReformDatetime` | 日時を `yyyy/MM/dd HH:mm:ss` 形式の文字列に整形 |
| `GetExportFolder` | エクスポート先の親フォルダパスを決定（未指定時は対象ファイルと同じ階層の `ExportedVBA` フォルダ） |
| `GetExportPath` | 対象ファイルごとのエクスポート先サブフォルダパスを決定（例: `Book1.xlsm` → `Book1_xlsm`） |
| `InitLog` / `EndLog` | ログの開始・終了 |
| `LogError` / `LogInfo` / `LogDebug` | ログレベル別の出力（`LogInfo`はレベル2以上、`LogDebug`はレベル9以上で出力） |
| `_LogFormat` | 実際のログ整形・出力処理（`Write-Host` でコンソールに出力） |

### 7-6. CommonWindowLog.ps1（SVN連携時のログ表示用）

`Common.ps1` の `_LogFormat` を**上書き（オーバーライド）**するファイルです。

- `initWindowLog`：一時フォルダにログファイルを作成し、**別のPowerShellウィンドウを新規に立ち上げて `Get-Content -Wait` でログをリアルタイム表示**する
- `_LogFormat`：ログをコンソールではなく、上記の一時ログファイルに追記する（`>> $script:log`）

SVNのコミット処理はTortoiseSVN内部で走るため、通常のコンソールが見えません。そこで**別ウィンドウを開いてログを流し込む**ことで、進捗をユーザーに見せる仕組みです。

---

