@echo off
rem Windows用アプリ（Ekonte\Ekonte.exe）を作る。Windows上で実行する。
rem Python 3.9〜3.12 をインストールしておくこと（https://www.python.org/）。
chcp 65001 > nul
cd /d %~dp0
set VENV=%USERPROFILE%\.venvs\ekonte
set BUILD=%USERPROFILE%\ekonte-build
if not exist "%VENV%\Scripts\python.exe" (
  py -3 -m venv "%VENV%" || python -m venv "%VENV%"
)
"%VENV%\Scripts\python.exe" -m pip install --upgrade pip
"%VENV%\Scripts\pip.exe" install -r requirements.txt pyinstaller==6.22.3 || exit /b 1
"%VENV%\Scripts\pyinstaller.exe" --noconfirm --clean --distpath "%BUILD%\dist" --workpath "%BUILD%\work" ekonte.spec || exit /b 1
echo 完成: %BUILD%\dist\Ekonte\Ekonte.exe
