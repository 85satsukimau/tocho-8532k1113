#!/bin/bash
cd "$(dirname "$0")"

# 1. Excelから「刀帳アプリ.html」を最新に更新
python3 update_data.py ../../刀剣乱舞データベース.xlsx 刀帳アプリ.html

# 2. GitHub用の「index.html」としてコピー
cp 刀帳アプリ.html index.html

# 3. GitHubへ自動アップロード（スマホへ反映）
git add index.html
git commit -m "データ自動更新"
git push -u origin main

echo "=========================================="
echo "スマホ（GitHub Pages）への反映が完了しました！"
echo "=========================================="
