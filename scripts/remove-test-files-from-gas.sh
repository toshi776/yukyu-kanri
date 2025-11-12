#!/bin/bash

# GAS上のtestファイルを削除するスクリプト
# testフォルダを一時的に退避させてから、clasp pushで同期する

set -e

echo "🗑️  GAS上のtestファイルを削除します..."

# 現在のディレクトリを保存
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

cd "$PROJECT_DIR"

# testフォルダが存在するか確認
if [ ! -d "test" ]; then
  echo "❌ testフォルダが見つかりません"
  exit 1
fi

echo "📦 testフォルダを一時退避中..."
if [ -d "test.backup" ]; then
  rm -rf test.backup
fi
mv test test.backup

# .claspignoreが正しく機能していることを確認
echo "📋 .claspignoreの確認..."
cat .claspignore

echo "📤 GASにプッシュ中（testファイルを削除）..."
clasp push --force

echo "📦 testフォルダを復元中..."
mv test.backup test

echo "✅ GAS上のtestファイルを削除しました"
echo ""
echo "🔄 新しいバージョンを作成してデプロイします..."

# 新しいバージョンを作成
VERSION=$(clasp version "Remove test files from GAS")
echo "✅ バージョン作成完了: $VERSION"

# バージョン番号を抽出（例: "Created version 48" から "48" を抽出）
VERSION_NUM=$(echo "$VERSION" | grep -oP '\d+$')

if [ -n "$VERSION_NUM" ]; then
  echo "📋 バージョン番号: $VERSION_NUM"

  # デプロイメントIDを取得
  DEPLOYMENT_ID=$(clasp deployments | grep -oP 'AKfycby[a-zA-Z0-9_-]+' | head -1)

  if [ -n "$DEPLOYMENT_ID" ]; then
    echo "🔄 Webアプリデプロイメントを更新中..."
    clasp deploy --deploymentId "$DEPLOYMENT_ID" --versionNumber "$VERSION_NUM" --description "Remove test files from GAS"
    echo "✅ デプロイ完了！"
    echo "📊 デプロイメントID: $DEPLOYMENT_ID"
    echo "📋 バージョン: $VERSION_NUM"
  else
    echo "⚠️  デプロイメントIDが見つかりませんでした"
  fi
else
  echo "⚠️  バージョン番号を抽出できませんでした"
fi

echo ""
echo "✅ すべての処理が完了しました"
