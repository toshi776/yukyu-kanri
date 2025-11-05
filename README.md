# 有給管理システム

Google Apps Script (GAS) + Googleスプレッドシートで構築した有給休暇管理システムです。

## 概要

事業所の有給休暇を効率的に管理するためのWebアプリケーションです。
- 利用者の有給残日数管理
- 有給申請・承認フロー
- 統計レポート機能
- 自動通知機能

## 開発環境のセットアップ

### 必要なツール
- Node.js
- clasp (Google Apps Script CLI)
- Python 3 (データ移行用)

### 初回セットアップ

1. リポジトリをクローン
```bash
git clone https://github.com/toshi776/yukyu-kanri.git
cd yukyu-kanri
```

2. claspで認証
```bash
clasp login
```

3. GASプロジェクトと同期
```bash
clasp pull
```

## 開発フロー

### コードの編集
ローカルでファイルを編集後、GASにプッシュ：
```bash
clasp push
```

### デプロイ
自動デプロイスクリプトを使用：
```bash
./deploy.sh
```

または手動でデプロイ：
```bash
clasp push
clasp create-version "バージョン説明"
clasp update-deployment [DEPLOYMENT_ID] --versionNumber [VERSION]
```

## ファイル構成

### メインファイル
- `Code.js` - メインロジック、Webアプリエントリーポイント
- `utils.js` - ユーティリティ関数
- `admin.html` - 管理画面UI
- `form.html` - 申請フォームUI
- `personal.html` - 個人用ページ

### 機能モジュール
- `leave-grant.js` - 有給付与機能
- `notification.js` - 通知システム
- `statistics-report.js` - 統計レポート生成
- `trigger-manager.js` - トリガー管理

### テスト
- `test-runner.js` - テストランナー
- `test-*.js` - 各機能のテストコード
- `system-integration-test.js` - 統合テスト

### その他
- `migrate_data.py` - Excelデータ移行スクリプト
- `deploy.sh` - 自動デプロイスクリプト
- `.clasp.json` - claspプロジェクト設定

## データ移行

Excelファイルからスプレッドシートへのデータ移行：
```bash
python3 migrate_data.py
```

## プロジェクト情報

### スクリプトID
`1JBx6pQKBVB68Ud2pbQf6n2YmLYHeSWuM_2vV2hD9zfGdI6BC1MQJOY-6`

### デプロイメントID（開発環境）
`AKfycbyJNqVvi4wYJwjXEc5Y9QF7qV-08M9uk4396sAo7Lu0i0lsY2RlCtbAPVMWaeYiKeKn`

## ライセンス

私的利用
