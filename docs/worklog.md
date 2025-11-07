# 作業ログ

---

## 2025-11-06: 新環境でのclasp認証問題を完全解決

### 背景
新しいノートPC（WSL2環境）で作業を開始したため、clasp認証情報（`.clasprc.json`）が存在せず、GASへのデプロイができない状態だった。昨日（11/5）も認証を試みたが、OAuth同意画面の設定不備やスコープ不足で失敗していた。

### 問題の詳細

#### 初期状態
- ✅ claspインストール済み（v2.5.0）
- ✅ `.clasp.json`（プロジェクト設定）は存在
- ❌ `.clasprc.json`（認証情報）が不在
- ❌ デプロイ実行不可

#### 遭遇したエラー

**1. 標準clasp loginでのエラー**
```
Error 400: invalid_request
アクセスをブロック: clasp – The Apps Script CLI のリクエストは無効です
```
→ GoogleがclaspのデフォルトOAuthクライアントIDを制限している

**2. カスタムOAuth認証情報でのエラー（初回）**
```
Error 403: access_denied
yukyu-kanri は Google の審査プロセスを完了していません
```
→ OAuth同意画面でテストユーザーが未設定

**3. スコープ不足エラー**
```
No access, refresh token, API key or refresh handler callback is set.
```
→ カスタム認証情報使用時、限定スコープ（`script.webapp.deploy`のみ）しか取得されず、`script.projects`などが不足

### 解決手順

#### ステップ1: クリーンアップ
- 前日作成した不完全なGCPプロジェクト `yukyu-kanri` を削除
- 認証情報をリセット

#### ステップ2: 新規GCPプロジェクト作成
- プロジェクト名: `yukyu-kanri`
- プロジェクトID: `polynomial-net-477401-k0`
- プロジェクト番号: `653762862710`

#### ステップ3: GASスクリプトとGCPプロジェクトを関連付け
```
GASエディタ → プロジェクトの設定 → GCPプロジェクトを変更
プロジェクト番号を入力: 653762862710
```

#### ステップ4: Apps Script API有効化
```
https://console.cloud.google.com/apis/library/script.googleapis.com?project=653762862710
→ 「有効にする」をクリック
```

#### ステップ5: OAuth同意画面とクライアントID作成

**OAuth同意画面の設定:**
- User Type: 外部
- アプリ名: `yukyu-kanri`
- テストユーザー追加: `toshi776@gmail.com` ← **重要！**
- または「アプリを公開」（個人利用の場合はこちらが簡単）

**OAuthクライアントID作成:**
- アプリケーション種類: デスクトップアプリ
- 名前: `clasp`
- JSONダウンロード

#### ステップ6: clasp認証

**カスタム認証での問題発覚:**
カスタムOAuth認証情報を使用すると、claspが限定的なスコープしか要求しない：
```bash
clasp login --creds client_secret_xxx.json
# → スコープ: script.webapp.deploy のみ（不十分）
```

**最終的な解決策:**
標準のclasp loginを再試行したところ、今回は成功！
```bash
clasp login
# → すべての必要なスコープを取得:
#   - script.deployments
#   - script.projects
#   - script.webapp.deploy
#   - drive.file
#   - その他7スコープ
```

**成功の理由:**
- GCPプロジェクトとの関連付けが正しく設定されていた
- Apps Script APIが有効化されていた
- OAuth同意画面が適切に設定されていた
→ これらの前準備により、標準clasp loginが正常動作

#### ステップ7: .clasp.json修正

プロジェクトIDを追加：
```json
{
  "scriptId": "1JBx6pQKBVB68Ud2pbQf6n2YmLYHeSWuM_2vV2hD9zfGdI6BC1MQJOY-6",
  "projectId": "polynomial-net-477401-k0",  // ← 追加
  "rootDir": ".",  // ← "" から "." に変更
  ...
}
```

#### ステップ8: deploy.sh修正

古いclaspコマンドを最新版に対応：
```bash
# 修正前（v1系コマンド）
clasp create-version "説明"
clasp update-deployment $ID --versionNumber $NUM

# 修正後（v2系コマンド）
clasp version "説明"
clasp deploy -i $ID -V $NUM -d "説明"
```

### 成功！

```bash
./deploy.sh
```

**結果:**
```
✅ バージョン作成完了: Created version 8.
✅ デプロイ完了！
🌐 Webアプリが更新されました
📊 デプロイメントID: AKfycbyJNqVvi4wYJwjXEc5Y9QF7qV-08M9uk4396sAo7Lu0i0lsY2RlCtbAPVMWaeYiKeKn
📋 バージョン: 8
```

プッシュされたファイル: 15ファイル
- admin.html, form.html, personal.html
- Code.js, utils.js
- leave-grant.js, notification.js, statistics-report.js, trigger-manager.js
- 各種テストファイル

### 学んだこと・ポイント

#### 1. clasp認証の正しい手順
新環境でのclasp認証は、以下の順序が重要：
1. GCPプロジェクト作成
2. GASスクリプトにGCPプロジェクトを関連付け
3. Apps Script API有効化
4. OAuth同意画面設定（テストユーザー追加 or 公開）
5. **標準clasp login**を実行

#### 2. カスタムOAuth認証の落とし穴
`clasp login --creds`はスコープが限定される問題がある。標準の`clasp login`の方が確実。

#### 3. OAuth同意画面のテストユーザー設定
「外部」アプリの場合、自分のメールアドレスをテストユーザーに追加するか、アプリを公開する必要がある。

#### 4. deploy.shの互換性
claspのバージョンアップにより、コマンド名が変更されている：
- `create-version` → `version`
- `update-deployment` → `deploy -i`

#### 5. .clasp.jsonの設定
- `rootDir: ""`は動作しない。`"."`を指定する
- `projectId`を追加すると、プロジェクト識別がスムーズ

### 次回同じ状況になった場合の手順書

新しいマシンでclasp環境をセットアップする場合：

```bash
# 1. Node.js/clasp確認
node --version
clasp --version

# 2. リポジトリクローン
git clone https://github.com/toshi776/yukyu-kanri.git
cd yukyu-kanri

# 3. 認証（GCP設定が既にある場合）
clasp login

# 4. 動作確認
clasp push --force

# 5. デプロイ
./deploy.sh
```

**初回セットアップの場合:**
上記「解決手順」のステップ2〜6を実施してから、手順3以降を実行。

### 成果物
- ✅ 新環境でのclasp認証完了
- ✅ GASへのプッシュ成功
- ✅ 自動デプロイスクリプト修正・動作確認
- ✅ バージョン8のデプロイ成功
- ✅ この詳細ドキュメント

---

## 2025-11-05: 初回認証試行（未完了）

### 作業概要
- Phase 1（個人URLアクセス制御）のコード確認
- `clasp` 認証を独自 OAuth クライアントで再設定
- 複数回のログイン試行を実施するも CLI 側の制約で未完了
- 報告書 `implementation-report.md` の Markdown 化
- 認証試行内容をドキュメント化

### 認証トラブル概要
- Google Cloud Console で新規 OAuth クライアント ID/Secret を作成（値は非公開）
- `appsscript.json` に必要スコープを追記
- `clasp login --creds` を複数回実行するも、トークン再読込で未ログイン状態
- `CLASP_SKIP_ENABLE_APIS` など環境変数を試すも改善せず
- 現状はブラウザの Apps Script エディタからの操作が必要

### 結果
→ 翌日（11/6）に完全解決

---

## 2025-11-07: 有給申請機能のテスト・バグ修正

### 背景
プロジェクトから長期間離れていたため、システムの実装内容と使い方を忘れていた。有給管理システムの動作確認を行い、複数の不具合を発見・修正。

### 発見した問題と修正内容

#### 1. 管理画面の申請管理タブが表示されない問題（解決✅）

**問題:**
- 管理画面で「申請承認」タブをクリックしても申請データが表示されない
- コンソールに`getApplications呼び出し`のログが出ない

**原因:**
- `admin.html`に`showTab()`関数が2箇所で定義されていた（360行目と1000行目）
- 後の方（1000行目）が優先され、`loadApplications()`の呼び出しが実行されなかった

**修正内容:**
- 360行目の古い`showTab()`関数を削除
- 1000行目（修正後976行目）の`showTab()`関数に`loadApplications()`の呼び出しを追加

**デプロイ:**
- Version 18/19をデプロイ

**結果:**
✅ 管理画面の申請管理タブが正常に動作するようになった

#### 2. 個人ページのリロード問題（未解決❌）

**問題:**
- 個人ページで有給申請を送信後、ページが白い画面で止まる
- データはサーバー側に正常に保存されている
- ページリロードが正常に完了しない

**試行した修正:**

**試行1 (v16-17): `.always()`メソッドの削除**
- 問題: `google.script.run`APIには`.always()`メソッドが存在しない
- 修正: ボタンのリセット処理を`withSuccessHandler`と`withFailureHandler`の両方に移動
- 結果: 問題解決せず

**試行2 (v20): `location.reload()`の変更**
- 問題: `location.reload()`の動作が不安定
- 修正: `window.location.href = window.location.href`に変更
- 結果: 問題解決せず

**試行3 (v21): デバッグログの追加**
- 目的: 処理フローを追跡
- 結果: ログは正常に表示されるが、リロード後に白い画面になる
```
📤 申請送信開始...
✅ サーバーからのレスポンス: {success: true, newRemaining: 4}
✅ 申請成功 - 新しい残日数: 4
🔄 2秒後にページをリロードします...
```

**試行4 (v22): 即座のリロード**
- 問題: `setTimeout`や`showAlert`が原因の可能性
- 修正: すべての中間処理をスキップし、即座に`window.location.href = window.location.href`を実行
- 結果: 問題解決せず。ログは「🔄 すぐにページをリロードします...」まで到達するが、その後白い画面になる

**現在の状態:**
- サーバー側の処理は正常（申請データは保存される）
- JavaScriptのリロード処理は実行される（ログから確認）
- リロード後のページレンダリングで問題が発生している可能性
- 根本原因は未特定

### デプロイ履歴

| Version | 説明 | 修正内容 |
|---------|------|----------|
| 16-17 | Fix personal page reload by removing .always() method | `.always()`削除 |
| 18-19 | Fix duplicate showTab function - add loadApplications() call | showTab重複修正 |
| 20 | Fix personal page reload using window.location.href | location.href変更 |
| 21 | Add debug logs to track submitRequest flow | デバッグログ追加 |
| 22 | Immediate page reload on success without alerts | 即座リロード |

### 使用しているデプロイメントID
- 主要: `AKfycbyJNqVvi4wYJwjXEc5Y9QF7qV-08M9uk4396sAo7Lu0i0lsY2RlCtbAPVMWaeYiKeKn` (v22)
- 副次: `AKfycbxWVjrZ1NfAiAXv0NnO8atjpHLmgKcFqGoBgn8OnfvfWCs3mmN_lwrb_we0w7A10P1R` (v19)

### 次回の対応方針

**個人ページリロード問題の調査項目:**
1. リロード後のページでJavaScriptエラーが発生していないか
2. リロード時のNetworkタブでHTTPエラーが発生していないか
3. `renderPersonalPage()`関数でエラーが発生していないか
4. 代替案: リロードではなく、DOMを直接更新する方法を検討

**確認すべきログ:**
- 白い画面になった後のConsoleエラー
- NetworkタブのHTTPステータス
- GASの実行ログ（リロード時のrenderPersonalPage呼び出し）

### 成果物
- ✅ 管理画面の申請管理タブ修正完了
- ❌ 個人ページのリロード問題は未解決（次回継続）
- 📝 詳細な調査ログを記録

---
