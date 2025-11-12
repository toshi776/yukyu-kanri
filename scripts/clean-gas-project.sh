#!/bin/bash

# GAS上のtestファイルを完全に削除するスクリプト
# srcフォルダとappsscript.jsonのみをプッシュする

set -e

echo "🗑️  GAS上の不要なファイルを削除します..."

# 現在のディレクトリを保存
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

cd "$PROJECT_DIR"

# 一時ディレクトリを作成
TMP_DIR=$(mktemp -d)
echo "📦 一時ディレクトリ: $TMP_DIR"

# srcとappsscript.jsonのみをコピー
echo "📋 必要なファイルのみをコピー中..."
cp -r src "$TMP_DIR/"
cp appsscript.json "$TMP_DIR/"
cp .clasp.json "$TMP_DIR/"

# 現在のディレクトリを保存
ORIGINAL_DIR=$(pwd)

# 一時ディレクトリに移動
cd "$TMP_DIR"

echo "📤 GASにプッシュ中（src配下のみ）..."
clasp push --force

# 元のディレクトリに戻る
cd "$ORIGINAL_DIR"

# 一時ディレクトリを削除
rm -rf "$TMP_DIR"

echo "✅ GAS上のtestファイルを削除しました"
echo ""
echo "🔄 新しいバージョンを作成してデプロイします..."

# 新しいバージョンを作成
VERSION=$(clasp version "Clean GAS project - remove test files")
echo "✅ バージョン作成完了: $VERSION"

# バージョン番号を抽出（例: "Created version 49" から "49" を抽出）
VERSION_NUM=$(echo "$VERSION" | grep -oP '\d+$')

if [ -n "$VERSION_NUM" ]; then
  echo "📋 バージョン番号: $VERSION_NUM"

  # デプロイメントIDを取得
  DEPLOYMENT_ID=$(clasp deployments | grep -oP 'AKfycby[a-zA-Z0-9_-]+' | head -1)

  if [ -n "$DEPLOYMENT_ID" ]; then
    echo "🔄 Webアプリデプロイメントを更新中..."
    clasp deploy --deploymentId "$DEPLOYMENT_ID" --versionNumber "$VERSION_NUM" --description "Clean GAS project - remove test files"
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
echo ""
echo "📋 GAS上のファイル一覧:"
clasp status
