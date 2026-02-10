# ソフトウェア仕様書
# PLC連携型 Excel PDF自動変換ツール（16シート対応・3画面構成）
# Software Specification v2.0

---

## 1. システム概要

三菱電機製PLC（Q/iQ-R/iQ-Fシリーズ）からMCプロトコルで16ビット指令（Dデバイス）を受信し、
指定Excel内の該当シートをPDF化するWindows常駐型GUIアプリケーション。

### 基本スペック

| 項目 | 値 |
|---|---|
| 言語 | Python 3.10+ |
| GUI | tkinter (標準ライブラリ) |
| PLC通信 | pymcprotocol (MCプロトコル Type3E) |
| Excel操作 | pywin32 (COM: Excel.Application) |
| 配布形態 | PyInstaller による単一EXE |
| 動作OS | Windows 10/11 (Microsoft Excel必須) |
| タイムゾーン | JST (UTC+9) 固定 |

### 依存ライブラリ (requirements.txt)

```
pywin32>=306
pymcprotocol>=0.4.0
pyinstaller>=6.0
```

---

## 2. ファイル構成

```
project/
├── config.py                # 設定管理 (JSON永続化、パスワード)
├── plc_comm.py              # PLC通信 (実機接続 + MockPLC)
├── core.py                  # コアロジック (変換・ファイル管理・制御)
├── excel_pdf_processor.py   # GUIアプリケーション (3画面)
├── tests/
│   ├── __init__.py
│   └── test_processor.py    # ユニットテスト (61テスト)
├── requirements.txt
├── build.spec               # PyInstaller設定
├── build.bat                # ビルドスクリプト
└── plc_pdf_config.json      # 設定ファイル (実行時自動生成)
```

---

## 3. 画面構成（3画面）

### 3.1 画面遷移図

```
┌─────────────┐   パスワード認証   ┌─────────────┐
│  動作画面    │ ─────────────────→ │  設定画面    │
│ (Operation)  │ ←───────────────── │ (Settings)   │
│              │  「保存して戻る」   │              │
│  ※起動時     │  または「キャンセル」│              │
│  デフォルト   │                    └─────────────┘
│              │
│              │   認証不要          ┌─────────────┐
│              │ ─────────────────→ │ テストモード  │
│              │ ←───────────────── │ (Test Mode)  │
└─────────────┘  「動作画面に戻る」  └─────────────┘
```

### 3.2 動作画面（Operation Screen）

**起動時のデフォルト画面。自動起動有効時はこの画面で自動的に監視開始する。**

#### レイアウト

```
┌──────────────────────────────────────────────────────┐
│ [動作画面 / Operation]            [設定画面] [テストモード] │
├──────────────────────────────────────────────────────┤
│ ステータス / Status                                     │
│ Heartbeat: ON/OFF  Ready: ON/OFF  Busy: ON/OFF       │
│ Error: ON/OFF      PLC: 接続中/未接続  Files: N        │
├──────────────────────────────────────────────────────┤
│ 現在の設定 / Current Config (読み取り専用)              │
│ Input:   C:\work\input                                │
│ Output:  C:\work\output                               │
│ Archive: C:\work\archive                              │
│ PLC: 192.168.1.10:5000  Cmd=D0  Cmp=D1  Mon=D2      │
├──────────────────────────────────────────────────────┤
│ [監視開始 / Start]  [監視停止 / Stop]                   │
├──────────────────────────────────────────────────────┤
│ ログ / Log                                            │
│ ┌──────────────────────────────────────────────┐     │
│ │ 12:00:01 [INFO] System started.               │     │
│ │ 12:00:01 [INFO] Controller started ...        │     │
│ │ ...                                           │     │
│ └──────────────────────────────────────────────┘     │
│                                        [Clear Log]    │
└──────────────────────────────────────────────────────┘
```

#### 動作仕様

| 要素 | 仕様 |
|---|---|
| ステータス表示 | D2レジスタを500ms周期で読み取り、各ビットをラベルに反映 |
| 色分け | Heartbeat/Ready=緑, Busy=オレンジ, Error=赤, OFF=グレー |
| 設定表示 | 現在のconfigの値を読み取り専用で表示（編集不可） |
| 監視開始ボタン | PLCExcelController.start() を呼び出し。成功後ボタン無効化 |
| 監視停止ボタン | PLCExcelController.stop() を呼び出し。全ラベルをリセット |
| ログ | ScrolledText (Consolas 9pt)、スレッドセーフ、Clear機能あり |

### 3.3 設定画面（Settings Screen）

**パスワード認証後のみアクセス可能。**

#### レイアウト

```
┌──────────────────────────────────────────────────────┐
│ [設定画面 / Settings]                 [動作画面に戻る] │
├──────────────────────────────────────────────────────┤
│ フォルダ設定 / Folder Settings                         │
│ 作業中フォルダ (Input):  [___________________] [参照...] │
│ PDF出力先 (Output):      [___________________] [参照...] │
│ Excel保存用 (Archive):   [___________________] [参照...] │
├──────────────────────────────────────────────────────┤
│ PLC接続設定 / PLC Connection                           │
│ IP: [192.168.1.10]  Port: [5000]                      │
│ 指令(Cmd): [D0]  完了(Cmp): [D1]  モニタ(Mon): [D2]   │
├──────────────────────────────────────────────────────┤
│ 自動起動 / Auto Start                                  │
│ [✓] 起動時に自動で監視開始する                          │
├──────────────────────────────────────────────────────┤
│ パスワード変更 / Change Password                        │
│ 現在のPW: [****]  新しいPW: [****]  確認: [****] [変更] │
├──────────────────────────────────────────────────────┤
│ [設定を保存して戻る / Save & Return] [保存せず戻る]      │
└──────────────────────────────────────────────────────┘
```

#### 動作仕様

| 要素 | 仕様 |
|---|---|
| パスワード認証 | 動作画面から遷移時にモーダルダイアログでパスワード入力を要求 |
| デフォルトPW | `0000` (SHA-256ハッシュで保存、平文は保持しない) |
| PW変更 | 現在PW認証 → 新PW + 確認一致 → 即座にJSON保存 |
| フォルダ選択 | tkinter filedialog.askdirectory() で選択 |
| 保存 | 「保存して戻る」でJSON書き出し → 動作画面へ遷移 |
| キャンセル | 変更を破棄して動作画面へ遷移 |

### 3.4 テストモード画面（Test Mode Screen）

**PLC実機不要。MockPLCConnection（メモリ上のレジスタ）を使用。**

#### レイアウト

```
┌──────────────────────────────────────────────────────┐
│ [テストモード / Test Mode (PLC不要)]   [動作画面に戻る] │
├──────────────────────────────────────────────────────┤
│ モックPLCステータス / Mock PLC Status                   │
│ Heartbeat: --  Ready: --  Busy: --  Error: --         │
│ Mock: 停止/稼働中                                      │
├──────────────────────────────────────────────────────┤
│ D0 シートビット選択 / Sheet Bit Selection               │
│ [Bit0 ][Bit1 ][Bit2 ][Bit3 ][Bit4 ][Bit5 ][Bit6 ][Bit7 ] │
│ Sheet1 Sheet2 Sheet3 Sheet4 Sheet5 Sheet6 Sheet7 Sheet8│
│ [Bit8 ][Bit9 ][Bit10][Bit11][Bit12][Bit13][Bit14][Bit15] │
│ Sheet9 Sheet10 ...                           Sheet16  │
│                                                       │
│ [全選択] [全解除]                                       │
│ D0値 (hex): 0x0005                                     │
├──────────────────────────────────────────────────────┤
│ [テスト監視開始] [テスト停止] [指令送信 / Send D0 Command] │
├──────────────────────────────────────────────────────┤
│ テストログ / Test Log                                   │
│ ┌──────────────────────────────────────────────┐     │
│ └──────────────────────────────────────────────┘     │
│                                            [Clear]    │
└──────────────────────────────────────────────────────┘
```

#### 動作仕様

| 要素 | 仕様 |
|---|---|
| チェックボックス | 16個 (Bit0~15)、各チェックボックスの変更でD0値をリアルタイム更新 |
| D0値表示 | hex形式で表示 (`0x0005` など) |
| テスト監視開始 | MockPLCConnection生成 → PLCExcelController に注入 → start() |
| 指令送信 | チェックされたビットをD0に書き込み (MockPLC) |
| ステータス | 動作画面と同じ形式でMockPLCのD2を500ms周期で読み取り表示 |
| 戻る | テスト実行中なら自動停止してから動作画面へ遷移 |
| ログ | テスト専用の別ログウィジェット (動作画面のログとは独立) |

---

## 4. 設定ファイル仕様 (plc_pdf_config.json)

### スキーマ

```json
{
  "input_folder": "C:\\work\\input",
  "output_folder": "C:\\work\\output",
  "archive_folder": "C:\\work\\archive",
  "plc_ip": "192.168.1.10",
  "plc_port": 5000,
  "command_device": "D0",
  "complete_device": "D1",
  "monitor_device": "D2",
  "plc_poll_interval": 0.2,
  "heartbeat_interval": 0.5,
  "password_hash": "<SHA-256 hex digest>",
  "auto_start": true
}
```

### 各フィールド

| キー | 型 | デフォルト | 説明 |
|---|---|---|---|
| input_folder | string | "" | 作業中フォルダ（Excelファイル1個のみ許可） |
| output_folder | string | "" | PDF出力先フォルダ |
| archive_folder | string | "" | 処理済Excelの保管先 |
| plc_ip | string | "192.168.1.10" | PLC IPアドレス |
| plc_port | int | 5000 | PLC MCプロトコルポート |
| command_device | string | "D0" | 指令レジスタ名 |
| complete_device | string | "D1" | 完了通知レジスタ名 |
| monitor_device | string | "D2" | モニタレジスタ名 |
| plc_poll_interval | float | 0.2 | D0ポーリング周期 (秒) |
| heartbeat_interval | float | 0.5 | D2.0トグル周期 (秒) |
| password_hash | string | SHA256("0000") | 設定画面パスワードのハッシュ |
| auto_start | bool | true | 起動時自動監視開始フラグ |

### パスワード管理

```
保存形式: SHA-256(平文パスワード).hexdigest()
デフォルト: SHA-256("0000") = "9af15b336e6a9619928537df30b2e6a2376569fcf9d7e773eccede65606529a0"
検証方法: SHA-256(入力値) == 保存されたハッシュ
```

---

## 5. PLCデバイスマッピング（通信仕様）

### 5.1 指令デバイス D0（入力：PCが読む）

16ビットワード。各ビットがExcelのシート番号に対応。

| ビット | 対応シート | ビット値 |
|---|---|---|
| D0.0 | シート1 (index 0) | 0x0001 |
| D0.1 | シート2 (index 1) | 0x0002 |
| D0.2 | シート3 (index 2) | 0x0004 |
| ... | ... | ... |
| D0.14 | シート15 (index 14) | 0x4000 |
| D0.15 | シート16 (index 15) | 0x8000 |

**ビットデコードアルゴリズム（必須実装）：**

```python
def decode_sheet_bits(command_word: int) -> list[int]:
    indices = []
    for i in range(16):
        if command_word & (1 << i):
            indices.append(i)
    return indices
```

- 複数ビットONの場合、**若い番号から順次処理**
- D0 = 0x0000 のときは無処理（ポーリング継続）

### 5.2 完了デバイス D1（出力：PCが書く）

| ビット | 機能 | 仕様 |
|---|---|---|
| D1.1 (Bit 1) | 完了パルス | 変換＋移動が正常完了後、100ms間ON → OFF（ワンショット） |

**ワンショット実装：**

```python
def send_oneshot(plc, device, bit, duration_s=0.1):
    plc.set_bit(device, bit)
    time.sleep(duration_s)
    plc.clear_bit(device, bit)
```

### 5.3 モニタデバイス D2（出力：PCが書く）

| ビット | 名称 | 条件 |
|---|---|---|
| D2.0 (Bit 0) | Heartbeat | ソフト稼働中、0.5秒周期でトグル（ON/OFF交互） |
| D2.1 (Bit 1) | Ready | エラーなし、ファイル1個待機中、指令受付可能な時ON |
| D2.2 (Bit 2) | Busy | 変換・移動処理実行中にON |
| D2.3 (Bit 3) | Error | ファイル数異常（0個/2個以上）、通信異常、変換失敗時にON |

### 5.4 ビット操作方式

Dレジスタはワード単位（16ビット）でしかread/writeできないため、
個別ビット操作は **Read-Modify-Write** パターンで実装する：

```python
def set_bit(device, bit):
    val = read_word(device)
    val |= (1 << bit)
    write_word(device, val & 0xFFFF)

def clear_bit(device, bit):
    val = read_word(device)
    val &= ~(1 << bit) & 0xFFFF
    write_word(device, val & 0xFFFF)
```

全操作は **threading.Lock** で排他制御すること（ポーリングスレッドとHeartbeatスレッドが同一接続を共有するため）。

---

## 6. メイン処理アルゴリズム（詳細フロー）

### 6.1 全体ライフサイクル

```
起動
 │
 ├─ config.json 読み込み（なければデフォルト生成）
 ├─ GUI構築（3画面分のFrame作成）
 ├─ 動作画面を表示
 │
 ├─ auto_start=true かつ フォルダ設定済み？
 │   ├─ YES → 500ms後にPLC接続＋監視自動開始
 │   └─ NO  → ユーザーの手動操作待ち
 │
 └─ mainloop()
```

### 6.2 監視開始フロー

```
監視開始ボタン押下 or 自動起動
 │
 ├─ フォルダ3つとも設定済みか検証
 │   └─ 未設定 → 警告ダイアログ表示して中断
 │
 ├─ output_folder / archive_folder が存在しなければ自動作成
 │
 ├─ PLCExcelController 生成
 │   ├─ plc_connection引数あり → 注入されたPLC使用（テストモード）
 │   └─ plc_connection引数なし → PLCConnection(ip, port) を新規作成
 │
 ├─ PLC接続 (connect)
 ├─ D1, D2 を 0 にクリア
 ├─ Heartbeatスレッド開始 (D2.0 トグル, daemonスレッド)
 ├─ ポーリングスレッド開始 (daemonスレッド)
 └─ GUIステータス更新開始 (500ms周期 root.after)
```

### 6.3 ポーリングループ（daemonスレッド）

```
while not stop_event:
    │
    ├─ _update_ready_status()
    │   ├─ input_folder内のExcelファイル数カウント
    │   ├─ 1個 → D2.1=ON (Ready), D2.3=OFF
    │   ├─ 2個以上 → D2.1=OFF, D2.3=ON (Error)
    │   └─ 0個 → D2.1=OFF のみ
    │
    ├─ D0 = read_word(command_device)
    ├─ D0 ≠ 0 ?
    │   ├─ YES → _handle_command(D0)
    │   └─ NO  → 何もしない
    │
    └─ stop_event.wait(0.2秒)  ← ポーリング間隔
```

### 6.4 コマンド処理フロー（_handle_command）

```
_handle_command(command_word) が呼ばれる
 │
 │  ★ 処理順序は厳密に守ること
 │
 ├─ 1. D2.2 = ON (Busy)
 │
 ├─ 2. ファイル数検証
 │   ├─ 0個 or 2個以上 → FileCountError
 │   │   ├─ ログにエラー出力
 │   │   ├─ D2.3 = ON (Error)
 │   │   ├─ GUIにエラーポップアップ表示
 │   │   └─ → finally へ (Busy OFF)
 │   └─ 1個 → excel_path 取得
 │
 ├─ 3. ビットデコード
 │   └─ decode_sheet_bits(command_word) → [0, 2, 4, ...] のようなリスト
 │
 ├─ 4. PDF変換  ★★★ 最重要セクション ★★★
 │   │
 │   │  convert_sheets_to_pdf(excel_path, sheet_indices, output_dir)
 │   │
 │   ├─ pythoncom.CoInitialize()
 │   ├─ excel = DispatchEx("Excel.Application")
 │   ├─ excel.Visible = False
 │   ├─ wb = Workbooks.Open(path, ReadOnly=True)
 │   │
 │   ├─ for idx in sheet_indices:
 │   │   ├─ sheet_num = idx + 1  (COMは1ベース)
 │   │   ├─ sheet_num > sheet_count → skip (ログ警告)
 │   │   ├─ sheet = wb.Worksheets(sheet_num)
 │   │   ├─ pdf_name = "{元ファイル名}_{シート名}_{yyyyMMdd_HHmmss}.pdf"
 │   │   └─ sheet.ExportAsFixedFormat(0, pdf_path)  ← xlTypePDF=0
 │   │
 │   └─ finally:  ★★★ ここが最重要 ★★★
 │       ├─ wb.Close(False)
 │       ├─ excel.Quit()
 │       ├─ time.sleep(0.3)     ← OSのファイルハンドル解放待ち
 │       └─ pythoncom.CoUninitialize()
 │
 │  ※ この finally が完了するまでファイル移動してはいけない
 │
 ├─ 5. ファイル移動 (Excel Quit完了後に実行)
 │   ├─ 保存先: archive_folder/{元名}_{yyyyMMdd_HHmmss}.xlsx
 │   ├─ 同名存在時はカウンタ付加 ({元名}_{ts}_{1}.xlsx)
 │   └─ shutil.move(excel_path, dest_path)
 │
 ├─ 6. 完了通知
 │   ├─ D1.1 = ON
 │   ├─ time.sleep(0.1)  ← 100ms
 │   └─ D1.1 = OFF
 │
 ├─ 7. D2.3 = OFF (Error クリア)
 │
 └─ finally:
     └─ D2.2 = OFF (Busy OFF)
```

### 6.5 ★ Quit→Move の順序保証（最重要制約）

```
【絶対に守るべき順序】

  wb.Close(False)           ← 1. ワークブックを閉じる
  excel.Quit()              ← 2. Excelプロセスを終了
  time.sleep(0.3)           ← 3. OSがファイルハンドル解放するのを待つ
  pythoncom.CoUninitialize() ← 4. COM解放

  --- ここまで convert_sheets_to_pdf() の finally ブロック内 ---
  --- ここから呼び出し元に戻ってから実行 ---

  shutil.move(...)           ← 5. ファイル移動（ここで初めて触れる）

【理由】
  ExcelのCOMプロセスがファイルをロックしているため、
  Quit()前にmoveするとPermissionErrorが発生する。
  finally内でQuitを保証し、関数リターン後にのみmoveを呼ぶ設計にする。
```

---

## 7. モジュール別詳細設計

### 7.1 config.py – 設定管理

#### クラス: AppConfig (dataclass)

| メソッド | 引数 | 戻り値 | 説明 |
|---|---|---|---|
| verify_password | plain: str | bool | SHA-256比較でパスワード検証 |
| change_password | new_plain: str | None | ハッシュを更新 |
| save | path: str\|None | str | JSONファイルに書き出し。パスを返す |
| load (classmethod) | path: str\|None | AppConfig | JSONから読み込み。ファイルなければデフォルト返却 |

#### 関数: hash_password

```python
def hash_password(plain: str) -> str:
    return hashlib.sha256(plain.encode()).hexdigest()
```

#### 設計ポイント
- loadで未知のキーは無視する（将来の互換性）
- ファイルが存在しなければデフォルト値のインスタンスを返す

### 7.2 plc_comm.py – PLC通信

#### クラス: PLCConnection

| メソッド | 説明 |
|---|---|
| connect() | pymcprotocol.Type3E() で TCP接続 |
| disconnect() | close() して状態リセット |
| read_word(device) | batchread_wordunits で1ワード読取 |
| write_word(device, value) | batchwrite_wordunits で1ワード書込 |
| set_bit(device, bit) | Read-Modify-Write でビットセット |
| clear_bit(device, bit) | Read-Modify-Write でビットクリア |
| write_bits(device, set_list, clear_list) | 複数ビット同時操作 |

- 全メソッドは `threading.Lock` で排他制御
- pymcprotocol は late import（非Windows環境でのテスト対応）

#### クラス: MockPLCConnection

PLCConnectionと**完全に同じAPI**を持つインメモリ実装。

| 内部状態 | 説明 |
|---|---|
| _registers: dict[str, int] | デバイス名→16ビット値のマップ |
| _lock: threading.Lock | スレッドセーフ保証 |
| _connected: bool | 接続状態フラグ |

- disconnect()でレジスタを全クリア
- read_wordで未登録デバイスは0を返す
- 全write操作で `& 0xFFFF` マスク（16ビット保証）

#### クラス: HeartbeatThread

- daemonスレッドでD2.0を0.5秒周期でトグル
- stop()時にD2.0をOFFに戻す

#### 関数: send_oneshot

- 指定bitをON → sleep(duration) → OFF

### 7.3 core.py – コアロジック

#### ユーティリティ関数

| 関数 | 説明 |
|---|---|
| jst_now() | JST (UTC+9) の現在時刻を返す |
| sanitize_filename(name) | Windows禁止文字 `<>:"/\|?*` を `_` に置換 |
| list_excel_files(folder) | フォルダ内のExcelファイルパスのリストを返す。`~$`で始まるテンプファイルは除外 |
| validate_single_file(folder) | ファイル数が1でなければ FileCountError を投げる |
| decode_sheet_bits(word) | 16ビットワードからONビットの0ベースインデックスリストを返す |
| convert_sheets_to_pdf(...) | 指定シートをPDF変換。finally内でExcel Quit保証 |
| archive_excel_file(...) | タイムスタンプ付きで保存フォルダへ移動 |

#### 対応Excelファイル拡張子

```python
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".xlsb"}
```

#### PDF出力ファイル名規則

```
{元ファイルのstem}_{シート名}_{yyyyMMdd_HHmmss}.pdf
例: 見積書_Sheet1_20260210_143022.pdf
```

#### Excelアーカイブファイル名規則

```
{元ファイルのstem}_{yyyyMMdd_HHmmss}{元拡張子}
例: 見積書_20260210_143022.xlsx
衝突時: 見積書_20260210_143022_1.xlsx, _2.xlsx, ...
```

#### クラス: PLCExcelController

| コンストラクタ引数 | 型 | 説明 |
|---|---|---|
| config | AppConfig | 全設定値 |
| logger | logging.Logger | ログ出力先 |
| on_error | callable\|None | エラー時コールバック(str) |
| on_status | callable\|None | ステータス更新コールバック(dict) |
| plc_connection | any\|None | None=実機接続、指定=注入（テスト用） |

| 公開メソッド | 説明 |
|---|---|
| start() | PLC接続、Heartbeat開始、ポーリング開始 |
| stop() | 全スレッド停止、PLC切断 |
| is_running (property) | ポーリングスレッドが生存中か |

### 7.4 excel_pdf_processor.py – GUI

#### クラス: TextHandlerWidget (logging.Handler)

- ScrolledTextウィジェットへスレッドセーフにログ出力
- `widget.after(0, _append, msg)` で mainthread にスケジュール

#### クラス: PasswordDialog (simpledialog.Dialog)

- パスワード入力用モーダルダイアログ
- `show="*"` でマスク表示
- OKで `result_pw` にパスワード格納、キャンセルで `None`

#### クラス: ExcelPdfProcessorApp

- 3画面を `ttk.Frame` で構築し、`pack_forget()` / `pack()` で画面切替
- 動作画面: `_build_operation_screen()`
- 設定画面: `_build_settings_screen()`
- テスト画面: `_build_test_screen()`

---

## 8. スレッド構成

```
メインスレッド (tkinter mainloop)
 │
 ├─ GUIイベント処理
 ├─ root.after(500ms) でステータス表示更新
 │
 ├─ Heartbeatスレッド (daemon)
 │   └─ 0.5秒周期で D2.0 トグル
 │
 └─ ポーリングスレッド (daemon)
     └─ 0.2秒周期で D0 読取 → コマンド処理
         └─ 変換処理中は同スレッド内で同期実行
```

**スレッドセーフ要件：**
- PLCConnection の全メソッドは `threading.Lock` で保護
- GUIへの書き込みは `widget.after(0, callback)` 経由
- stop_event は `threading.Event` で安全に停止通知

---

## 9. エラーハンドリング

| エラー条件 | 動作 |
|---|---|
| 作業中フォルダにファイル0個 | D2.1=OFF (Ready OFF)。D0指令受信時はFileCountError→D2.3=ON+ポップアップ |
| 作業中フォルダにファイル2個以上 | D2.1=OFF, D2.3=ON (Error)。D0指令受信時はFileCountError→ポップアップ |
| シート番号がブック内に存在しない | ログに警告出力してスキップ（エラーにはしない） |
| Excel COMエラー | 例外キャッチ → D2.3=ON → ログ出力。finallyでQuit保証 |
| PLC通信エラー | ログに例外出力 → D2.3=ON。ポーリングは次サイクルで再試行 |
| Excel Quit前のファイルロック | finally内でQuit保証により防止。sleep(0.3)でハンドル解放待ち |

---

## 10. テスト仕様

### テストファイル: tests/test_processor.py

61テスト、以下のカテゴリ：

| テストクラス | テスト数 | テスト対象 |
|---|---|---|
| TestSanitizeFilename | 8 | ファイル名サニタイズ（日本語、禁止文字、空文字） |
| TestJstNow | 3 | JST タイムゾーン、オフセット、フォーマット |
| TestExcelExtensions | 2 | 対応/非対応拡張子 |
| TestDecodeSheetBits | 7 | ビットデコード（0ビット、1ビット、全ビット、ソート順） |
| TestListExcelFiles | 7 | フォルダ内ファイル列挙（空、テンプ除外、非Excel除外） |
| TestValidateSingleFile | 3 | 単一ファイル検証（0個エラー、1個OK、2個エラー） |
| TestArchiveExcelFile | 3 | ファイル移動（基本、ディレクトリ自動作成、タイムスタンプ） |
| TestAppConfig | 4 | JSON永続化、デフォルト値、未知キー無視 |
| TestAppConfigPassword | 6 | パスワード検証・変更・保存復元・ハッシュ |
| TestAppConfigAutoStart | 2 | 自動起動フラグのデフォルトと永続化 |
| TestMockPLCConnection | 8 | MockPLC全操作（接続、read/write、bit、16bitマスク） |
| TestPLCConnectionMock | 6 | 実PLCConnection のモック使用テスト |
| TestSendOneshot | 2 | ワンショットパルス（モック＋MockPLC） |

### テスト実行コマンド

```bash
python -m unittest tests.test_processor -v
```

### テスト方針
- COM (pywin32) と実PLC通信はモック使用
- MockPLCConnection は実オブジェクトとして直接テスト可能
- ファイル操作テストは `tempfile.mkdtemp()` を使用

---

## 11. ビルド・配布

### EXEビルド手順

```batch
pip install -r requirements.txt
pyinstaller build.spec
# → dist/PLCExcelPdfConverter.exe
```

### PyInstaller設定のポイント

| 設定 | 値 | 理由 |
|---|---|---|
| --onefile | 有効 | 単一EXE配布 |
| --windowed | 有効 | コンソール非表示 |
| hiddenimports | win32com, pymcprotocol, config, plc_comm, core | 動的importの検出漏れ防止 |

---

## 12. 実装上の注意事項（AIコーディング向け）

### 必須順守事項

1. **Quit→Move 順序**: Excel COM の `Quit()` は必ず `finally` 内で実行し、
   `shutil.move()` はその関数の return 後に呼ぶこと。**同一 try 内で move しないこと。**

2. **16ビットマスク**: 全 write 操作で `& 0xFFFF` を付けること。
   PLCのDレジスタは16ビットワード。

3. **スレッドセーフ**: PLC接続は1本のTCPソケットを共有するため、
   全 read/write を `threading.Lock` で保護すること。

4. **COM初期化**: `pythoncom.CoInitialize()` / `CoUninitialize()` は
   Excel操作を行うスレッドごとに呼ぶこと（ポーリングスレッド内で実行されるため）。

5. **テンプファイル除外**: `~$` で始まるExcelテンプファイルを
   ファイル列挙から除外すること。

6. **GUIスレッドセーフ**: ワーカースレッドからGUIを更新する場合は
   必ず `widget.after(0, callback)` を使うこと。

### 設計判断の根拠

| 判断 | 理由 |
|---|---|
| polling方式 (watchdog不使用) | PLC D0読取が主トリガのため、ファイル監視は補助的 |
| DispatchEx (DispatchではなくEx) | プロセス分離でExcel既存インスタンスの影響を排除 |
| ReadOnly=True | ファイルロック軽減 |
| sleep(0.3) after Quit | Windows OS のファイルハンドル解放ラグ対策 |
| SHA-256 for password | 平文保存回避（saltなしだが、ローカル産業用途では十分） |
| dataclass + asdict | JSON変換の自動化 |
| late import (pymcprotocol, win32com) | 非Windows環境でのテスト実行を可能にする |
