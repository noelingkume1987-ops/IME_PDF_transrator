# Excel to PDF Individual Sheet Processor (Pro Edition)

Excelファイルの各シートを個別のPDFに自動変換するWindows用ツールです。

## Features / 機能

- **GUI設定画面** - フォルダ選択、監視スパン設定、ステータス表示
- **自動監視** - 指定フォルダをポーリングし、新規・更新Excelファイルを検知
- **シート別PDF変換** - 各シートを個別のPDFファイルとして出力
- **タイムスタンプ命名** - `{ファイル名}_{シート名}_{YYYYMMDD_HHMMSS}.pdf`
- **アーカイブ処理** - 変換済みExcelファイルをArchiveフォルダへ自動移動
- **EXE配布** - PyInstallerによるスタンドアロン実行ファイル

## Requirements / 動作要件

- Windows 10/11
- Microsoft Excel がインストール済みであること

### For Development / 開発環境

- Python 3.10+
- Dependencies: `pip install -r requirements.txt`

## Usage / 使い方

### EXE版

1. `ExcelPdfProcessorPro.exe` をダブルクリック
2. **監視フォルダ** を「Browse...」ボタンで選択
3. **出力先フォルダ** を「Browse...」ボタンで選択
4. **監視スパン** をスライダーまたは数値で設定（1～36,000秒）
5. 「**Start Monitoring / 監視開始**」をクリック
6. 監視フォルダにExcelファイルを配置すると自動でPDF変換が実行されます

### Python版

```bash
pip install -r requirements.txt
python excel_pdf_processor.py
```

## Build / ビルド方法

### Windows バッチファイル

```cmd
build.bat
```

### 手動ビルド

```cmd
pip install -r requirements.txt
pyinstaller build.spec
```

出力先: `dist/ExcelPdfProcessorPro.exe`

## Naming Rule / 命名規則

```
{元ファイル名}_{シート名}_{YYYYMMDD_HHMMSS}.pdf
```

Example:
```
見積書_A社_20260123_210005.pdf
見積書_B社_20260123_210005.pdf
```

## Architecture / 構成

| Component | Technology |
|---|---|
| GUI | Tkinter |
| Excel conversion | pywin32 (Excel.Application COM API) |
| File monitoring | Polling-based scanner |
| Packaging | PyInstaller (`--onefile`) |
| Timestamps | datetime (JST / UTC+9) |

## Notes / 注意事項

- 実行環境にMicrosoft Excelが必要です（win32com経由で操作するため）
- 監視対象フォルダへの読込・書込権限が必要です
- `~$` で始まるExcelの一時ファイルは自動的にスキップされます
- 処理済みファイルは監視フォルダ内の `Archive` サブフォルダに移動されます
