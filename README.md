# 有給管理システム

Google Apps Script (GAS) + Googleスプレッドシートで構築した有給休暇管理システムです。

## 概要

事業所の有給休暇を効率的に管理するためのWebアプリケーションです。

### 主な機能

#### コア機能
- **有給申請**: 利用者専用URLから簡単に申請（1日/半日対応）
- **承認/却下**: 管理画面またはメール内ボタンで即座に処理
- **申請取消**: 前日までの申請取消が可能
- **有給付与**: 入社日・勤続年数に応じた自動付与
- **有給失効**: 付与日から2年後の自動失効処理

#### 通知機能
- 申請時の承認者への自動通知
- ワンクリック承認リンク（メール内のボタンで承認/却下）
- 承認結果の申請者への通知
- 失効予告通知（3ヶ月前・1ヶ月前）

#### セキュリティ
- 個人専用URL（32文字ランダムキー）
- 承認トークンによる不正操作防止
- 重複申請防止

#### UI/UX
- レスポンシブデザイン（スマホ・タブレット・PC対応）
- リアルタイム画面更新
- 明確なエラーメッセージ

### システムステータス
- **現在のバージョン**: v61（2025-11-12）
- **ステータス**: 安定版・本番運用可能
- **テスト**: 基本機能の動作確認完了

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

## 使い方

### 利用者（有給申請）
1. 個人専用URLにアクセス（管理者から提供されたURL）
2. 申請日と日数を選択
3. 「有給申請を送信」ボタンをクリック
4. 承認結果をメールで確認

### 管理者（承認処理）
**方法1: メールから承認**
1. 申請通知メールを受信
2. メール内の「✓ 承認する」または「✗ 却下する」ボタンをクリック
3. 処理完了

**方法2: 管理画面から承認**
1. 管理画面URL `?admin=true` にアクセス
2. 「申請管理」タブで申請一覧を確認
3. 承認または却下ボタンをクリック

## 開発フロー

### コードの編集
ローカルでファイルを編集後、GASにプッシュ：
```bash
clasp push
```

### デプロイ
自動デプロイスクリプトを使用：
```bash
./scripts/deploy.sh
```

または手動でデプロイ：
```bash
clasp push
clasp deploy --versionNumber [VERSION] --description "説明"
```

## ファイル構成

```
yukyu-kanri/
├── src/                          # Google Apps Script ソースコード
│   ├── Code.js                   # メインコード、doGet/doPost
│   ├── utils.js                  # ユーティリティ、データ操作
│   ├── notification.js           # 通知・メール機能
│   ├── leave-grant.js            # 有給付与ロジック
│   ├── leave-grant-core.js       # 付与・失効コア処理
│   ├── statistics-report.js      # 統計レポート
│   ├── trigger-manager.js        # トリガー管理
│   ├── admin.html                # 管理画面UI
│   ├── personal.html             # 利用者専用画面UI
│   └── form.html                 # 申請フォーム（レガシー）
│
├── test/                         # テストコード
│   ├── test-runner.js
│   ├── test-annual-grant.js
│   ├── test-notification-production.js
│   ├── test-statistics-report.js
│   └── system-integration-test.js
│
├── scripts/                      # デプロイメント・移行スクリプト
│   ├── deploy.sh                 # 自動デプロイスクリプト
│   └── migrate_data.py           # Excelデータ移行スクリプト
│
├── docs/                         # ドキュメント
│   ├── worklog.md                # 作業ログ
│   ├── implementation-report.md  # 実装レポート
│   ├── auth-experiments.md       # 認証実験記録
│   └── kihonsekkei.txt           # 基本設計
│
├── .clasp.json                   # clasp設定
├── appsscript.json              # GAS設定
├── CLAUDE.md                    # プロジェクト指示書
└── README.md                    # このファイル
```

## データ移行

Excelファイルからスプレッドシートへのデータ移行：
```bash
python3 scripts/migrate_data.py
```

## 本番運用に向けて

### 現在の設定（テスト期間中）
- 承認者メール: `toshi776@gmail.com`（全事業所統一）
- 通知機能: 有効
- テストモード: 無効

### 本番運用時に変更が必要な項目

#### 1. 承認者メールアドレス設定
**ファイル**: `src/notification.js`（519-539行目）

現在はテスト用に全て `toshi776@gmail.com` に送信されています。
本番運用時は事業所別の承認者設定に変更してください。

```javascript
// テスト期間中（現在）
return 'toshi776@gmail.com';

// 本番運用時（コメントを外す）
var approvers = {
  'R': 'rise-manager@company.com',
  'P': 'paron-manager@company.com',
  'S': 'ciel-manager@company.com',
  'E': 'ebisu-manager@company.com'
};
var division = userId.substring(0, 1);
return approvers[division] || 'hr@company.com';
```

#### 2. 申請者メールアドレス設定
**ファイル**: `src/notification.js`（544-554行目）

現在は仮のメールアドレスが使用されています。
マスターシートにメールアドレス列を追加し、`getEmployeeEmail()`関数を実装してください。

### 動作確認チェックリスト
- [ ] 利用者専用URLでアクセスできるか
- [ ] 有給申請が正常に送信されるか
- [ ] 承認者にメール通知が届くか
- [ ] ワンクリック承認リンクが動作するか
- [ ] 申請者に承認結果メールが届くか
- [ ] 管理画面で申請一覧が表示されるか
- [ ] スマホで正常に表示されるか

## プロジェクト情報

### バージョン
- **現在**: v61（2025-11-12）
- **ステータス**: 安定版・本番運用可能

### GAS プロジェクト
- **スクリプトID**: `1JBx6pQKBVB68Ud2pbQf6n2YmLYHeSWuM_2vV2hD9zfGdI6BC1MQJOY-6`
- **デプロイメントID**: `AKfycbyJNqVvi4wYJwjXEc5Y9QF7qV-08M9uk4396sAo7Lu0i0lsY2RlCtbAPVMWaeYiKeKn`
- **スプレッドシートID**: `1ENEljNya5MPY2Z7FPZR0LAguFPeCunwbJ3B9kgSjyK0`

### URL
- **管理画面**: `https://script.google.com/macros/s/[DEPLOYMENT_ID]/exec?admin=true`
- **利用者画面**: `https://script.google.com/macros/s/[DEPLOYMENT_ID]/exec?key=[USER_KEY]`

## トラブルシューティング

### よくある問題

**Q: ワンクリック承認リンクをクリックしてもエラーになる**
- A: リンクの有効期限が切れているか、既に処理済みの可能性があります。管理画面から再度確認してください。

**Q: スマホで取り消しボタンが表示されない**
- A: v61で解決済みです。ブラウザのキャッシュをクリアしてください。

**Q: 申請後に残日数が更新されない**
- A: ページをリロードしてください。申請は正常に処理されています。

### サポート
詳細は `docs/worklog.md` を参照してください。

## ライセンス

私的利用
