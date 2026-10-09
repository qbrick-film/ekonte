#!/bin/bash
# Mac用アプリ（Ekonte.app）を作る。
# 注意: 書類フォルダが iCloud 同期されていると、iCloud が「.」で始まるフォルダの中身に非表示フラグを付け、
#       Qt が部品を読めなくなる。そのため Python 環境とビルド作業は iCloud の外（ホーム直下）で行う。
set -euo pipefail
cd "$(dirname "$0")"
VENV="$HOME/.venvs/ekonte"
BUILD="$HOME/ekonte-build"
if [ ! -x "$VENV/bin/python" ]; then
  python3 -m venv "$VENV"
  "$VENV/bin/pip" install --upgrade pip
fi
"$VENV/bin/pip" install -r requirements.txt pyinstaller==6.22.3
"$VENV/bin/pyinstaller" --noconfirm --clean --distpath "$BUILD/dist" --workpath "$BUILD/work" ekonte.spec
chflags -R nohidden "$BUILD/dist/Ekonte.app"
echo "完成: $BUILD/dist/Ekonte.app"
# アプリケーションフォルダに入れてあれば、新しい版に入れ替える（Dock の登録はそのまま使える）
if [ -d "/Applications/Ekonte.app" ]; then
  if pgrep -f "/Applications/Ekonte.app/" >/dev/null; then
    echo "※ アプリが起動中のため /Applications は更新していません。アプリを終了してからもう一度実行してください"
  else
    rm -rf "/Applications/Ekonte.app"
    ditto "$BUILD/dist/Ekonte.app" "/Applications/Ekonte.app"
    echo "更新: /Applications/Ekonte.app"
  fi
fi
