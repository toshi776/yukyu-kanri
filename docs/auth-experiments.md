# clasp 認証試行ログ

## 概要
- 目的: Codex 環境で `clasp push` を可能にするための OAuth 認証
- 実施日: 2025-11-05

## 実施した対策
1. Google Cloud Console で新規プロジェクトを作成し、Apps Script API を有効化
2. OAuth 同意画面を設定し、自分のアカウントをテストユーザーとして登録
3. OAuth クライアント ID/Secret を作成（値はリポジトリには保存しない）
4. `appsscript.json` に必要スコープを追記し、`clasp login --creds oauth-client.json --no-localhost` を実行
5. `CLASP_SKIP_ENABLE_APIS` などの環境変数を使って再試行
6. 認証成功メッセージは出るものの、`clasp login --status` は "You are not logged in" のまま

## 判明したこと
- トークンは `.clasprc.json` に保存されるが、再読込時に認証エラーとなる
- CLI のバージョンや環境差異で発生している可能性があり、即時の解決は困難
- 秘密情報（OAuth クライアント ID/Secret）はリポジトリに含めない方針とした

## 今後の方針
- ブラウザの Apps Script エディタを使ってテスト・デプロイを実施
- もしくは、自宅 PC で取得済みの `.clasprc.json` を安全に流用できるか検討
- CLI 以外のデプロイフロー（GitHub Actions + clasp など）も調査対象
