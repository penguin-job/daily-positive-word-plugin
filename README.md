# Daily Positive Word System

CSVデータをもとに、日替わりでポジティブなメッセージを表示するシステムです。
データはExcel VBAで前処理を行い、WordPressのカスタムプラグインで表示しています。

---

## システム構成図

構成図（PDF）はこちら
→ https://github.com/penguin-job/daily-positive-word-plugin/blob/main/sys-01_portfolio.pdf

---

## 概要

本プロジェクトでは、日々のポジティブな言葉を管理・表示するシンプルな仕組みを構築しました。
CSVデータをExcel VBAで整形・加工し、そのデータをWordPressプラグインで読み込み、日替わりでメッセージを表示します。

---

## システムの流れ

```
CSVデータ  
↓  
Excel VBA（重複チェック・整形）  
↓  
加工済みCSV  
↓  
WordPressプラグイン  
↓  
Webサイトに日替わり表示  
```

---

## 主な機能

* 日替わりでメッセージを表示
* CSVによるデータ管理
* Excel VBAによる前処理
* WordPressショートコード対応
* 軽量なカスタムプラグイン

---

## リポジトリ構成

```
daily-positive-word-plugin/
├ excel-vba/
│   ├ mod01_Main.bas            # メイン処理
│   ├ mod10_UI_Control.bas      # UI制御
│   ├ mod20_IO_File.bas         # ファイル入出力
│   ├ mod30_Biz_Process.bas     # データ加工ロジック
│   └ mod90_Util.bas            # 共通ユーティリティ
├ sample-data/
│   └ quotes.csv                # ポジティブワードデータ
├ wordpress-plugin/
│   └ daily-positive-word.php   # WordPressプラグイン本体
├ README.md
└ sys-01_portfolio.pdf          # システム構成図
```

