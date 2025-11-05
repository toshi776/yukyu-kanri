#!/bin/bash

# 有給管理システム 自動デプロイスクリプト
echo "🚀 有給管理システムのデプロイを開始します..."

# WebアプリのデプロイメントID
DEPLOYMENT_ID="AKfycbxCrhFyMBOjKM07JBnTYQg6rGJlPTfgOV8FQ2tQe6F7Tbtn2I2Bcpz4ZVLlsjJpX-PS"

# 現在のブランチを確認
BRANCH=$(git branch --show-current 2>/dev/null || echo "not a git repo")
echo "📌 現在のブランチ: $BRANCH"

# 変更をコミット（gitリポジトリの場合）
if [ "$BRANCH" != "not a git repo" ]; then
    echo "💾 変更を保存中..."
    git add -A
    git commit -m "Deploy: $(date '+%Y-%m-%d %H:%M:%S')" || echo "変更なし"
fi

# GASにプッシュ
echo "📤 GASへプッシュ中..."
clasp push --force

# 新しいバージョンを作成
echo "🔄 新しいバージョンを作成中..."
VERSION=$(clasp create-version "Auto deploy $(date '+%Y-%m-%d %H:%M:%S')")
echo "✅ バージョン作成完了: $VERSION"

# バージョン番号を抽出
VERSION_NUM=$(echo "$VERSION" | grep -o '[0-9]*$')

if [ -z "$VERSION_NUM" ]; then
    echo "❌ バージョン番号を取得できませんでした"
    exit 1
fi

echo "📋 バージョン番号: $VERSION_NUM"

# 既存のWebアプリデプロイメントを更新
echo "🔄 Webアプリデプロイメントを更新中..."
clasp update-deployment "$DEPLOYMENT_ID" --versionNumber "$VERSION_NUM" --description "Auto deploy $(date '+%Y-%m-%d %H:%M:%S')"

if [ $? -eq 0 ]; then
    echo "✅ デプロイ完了！"
    echo "🌐 Webアプリが更新されました"
    echo "📊 デプロイメントID: $DEPLOYMENT_ID"
    echo "📋 バージョン: $VERSION_NUM"
else
    echo "❌ デプロイメントの更新に失敗しました"
    echo "💡 GASエディタで手動更新が必要な場合があります"
    clasp open-script
    exit 1
fi