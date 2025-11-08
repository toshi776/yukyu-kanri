# 有給管理システム データ移行プロジェクト

## プロジェクト概要
エクセルデータ配下の有給管理ファイルをスプレッドシートに移行する暫定システム構築

## 完了した作業

### 1. データ分析 ✅
- 4つの事業所（ライズ、パロン、シエル、EBISU）
- 合計153名の有給データファイル分析完了
- 退職者フォルダは移行対象外として確認

### 2. 移行方針決定 ✅
**利用者番号採番ルール:**
- ライズ: R + 2桁番号（R01-R09）
- パロン: P + 2桁番号（P01-P51、欠番あり）
- シエル: S + 2桁番号（S01-S51）
- EBISU: E + 2桁番号（E01-E45）

**データ項目:**
- 利用者番号：事業所コード + ファイル番号
- 利用者名：Excelファイル名から抽出
- 残有給日数：管理簿タブのJ列最下行
- 備考：空欄

### 3. データ移行実行 ✅
**移行スクリプト:** `migrate_data.py`
- Python環境セットアップ完了（pandas, openpyxl）
- 各ExcelファイルからJ列データ抽出機能実装
- 管理簿タブ自動検出機能付き

**移行結果:**
- ライズ: 9名（多くが0日、1名のみ24日）
- パロン: 48名（8-40日の幅広い分布）
- シエル: 51名（前半34日、後半10-25日）
- EBISU: 45名（0-21日の分布）

**出力ファイル:** `有給管理_移行データ.csv`

## 次回作業時の継続ポイント

### すぐに使用可能なファイル
1. `migrate_data.py` - データ移行スクリプト
2. `有給管理_移行データ.csv` - 完成したCSVデータ（153名分）

### 次のステップ候補
1. GoogleスプレッドシートへのCSVインポート手順作成
2. 管理者向けメンテナンス機能追加
3. データ検証・チェック機能の実装
4. 定期更新スクリプトの作成

### 重要な技術情報
- Python環境: pandas, openpyxl インストール済み
- Excelファイル構造: 管理簿タブのJ列に残日数格納
- ファイル命名規則: `[番号].[氏名].xlsx`

## ディレクトリ構造
```
yukyu-kanri/
├── src/                          # Google Apps Script ソースコード
│   ├── Code.js                   # メインコード
│   ├── leave-grant.js            # 有給付与機能
│   ├── notification.js           # 通知機能
│   ├── statistics-report.js      # 統計レポート
│   ├── trigger-manager.js        # トリガー管理
│   ├── utils.js                  # ユーティリティ
│   ├── admin.html                # 管理画面
│   ├── form.html                 # 申請フォーム
│   └── personal.html             # 個人画面
├── test/                         # テストコード
│   ├── test-annual-grant.js
│   ├── test-notification-production.js
│   ├── test-runner.js
│   ├── test-statistics-report.js
│   └── system-integration-test.js
├── docs/                         # ドキュメント
│   ├── implementation-report.md
│   ├── auth-experiments.md
│   ├── kihonsekkei.txt
│   └── worklog.md
├── scripts/                      # デプロイメント・移行スクリプト
│   ├── deploy.sh
│   └── migrate_data.py
├── .clasp.json                   # clasp設定
├── .gitignore
├── AGENTS.md
├── appsscript.json              # GAS設定
├── CLAUDE.md                    # このファイル
└── README.md
```

## 実行コマンド
```bash
# データ移行実行
python3 scripts/migrate_data.py

# GASデプロイ
./scripts/deploy.sh

# clasp操作
clasp push                       # コードをGASにアップロード
clasp open                       # GASエディタを開く
clasp deploy                     # 新しいバージョンをデプロイ
```

## 作業実施ガイドライン

### リソース制限対策
大規模な実装作業を行う際は、以下のルールに従ってください：

1. **定期的な作業ログ記録**
   - 各機能の実装完了ごとに `docs/worklog.md` に進捗を記録
   - リソース制限で中断した場合でも、次回作業時に継続できるようにする
   - 最低限、以下の情報を記録：
     - 完了した作業内容
     - 実装したファイルと主要な変更点
     - デプロイバージョン番号
     - 未完了の残作業

2. **段階的なデプロイ**
   - 機能単位でデプロイを実行
   - 各デプロイ後に動作確認
   - バージョン管理を徹底

3. **作業の分割**
   - 大きな機能は小さなタスクに分割
   - TodoWriteで進捗を可視化
   - 1つのタスクが完了したら、必ずログに記録してから次へ

4. **ログ記録のタイミング**
   - 機能実装完了時
   - デプロイ成功時
   - エラー発生時
   - 作業を中断する前