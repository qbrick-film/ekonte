#!/bin/bash
# Mac用アプリ（Ekonte.app）を作る。
# 注意: 書類フォルダが iCloud 同期されていると、iCloud が「.」で始まるフォルダの中身に非表示フラグを付け、
#       Qt が部品を読めなくなる。そのため Python 環境とビルド作業は iCloud の外（ホーム直下）で行う。
set -euo pipefail
cd "$(dirname "$0")"
VENV="$HOME/.venvs/ekonte"
BUILD="$HOME/ekonte-build"
# Python 3.10 以上が必要（Pillow 12 のため。公式ビルドは 3.11）。Mac に最初から入っている python3 は 3.9 なので、
# uv（~/.local/bin/uv）で入れた 3.11 を優先して使う。無ければ PATH にある 3.10 以上を探す。
NEW_ENOUGH='import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)'
if [ -x "$VENV/bin/python" ] && ! "$VENV/bin/python" -c "$NEW_ENOUGH"; then
  echo "Python 環境が古いので作り直します（$("$VENV/bin/python" --version)）"
  rm -rf "$VENV"
fi
if [ ! -x "$VENV/bin/python" ]; then
  PY=""
  for c in "$("$HOME/.local/bin/uv" python find 3.11 2>/dev/null || true)" \
           "$(command -v python3.12 || true)" "$(command -v python3.11 || true)" "$(command -v python3.10 || true)" "$(command -v python3 || true)"; do
    if [ -n "$c" ] && [ -x "$c" ] && "$c" -c "$NEW_ENOUGH" 2>/dev/null; then PY="$c"; break; fi
  done
  if [ -z "$PY" ]; then
    echo "Python 3.10 以上が見つかりません。次のどちらかで入れてから、もう一度実行してください:"
    echo "  ~/.local/bin/uv python install 3.11   （uv がある場合）"
    echo "  https://www.python.org/downloads/ から Python 3.11 をインストール"
    exit 1
  fi
  "$PY" -m venv "$VENV"
  "$VENV/bin/pip" install --upgrade pip
fi
"$VENV/bin/pip" install -r requirements.txt pyinstaller==6.22.3 "setuptools>=83"
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
