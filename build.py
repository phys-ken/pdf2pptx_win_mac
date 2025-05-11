#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PDF to PPTX Converter用のビルドスクリプト
PyInstallerを使用して実行ファイル化します
"""
import os
import sys
import subprocess
import shutil
import platform

def main():
    """メインビルド処理"""
    print("PDF to PPTX Converterのビルドを開始します...")
    
    # ビルド設定
    entry_point = os.path.join('src', 'pdf2pptx_gui.py')
    app_name = "pdf2pptx_converter"
    icon_file = os.path.join('resources', 'app_icon.ico') if os.path.exists(os.path.join('resources', 'app_icon.ico')) else None
    
    # OSによって実行ファイルの拡張子を変更
    ext = '.exe' if platform.system() == 'Windows' else ''
    
    # PyInstallerのコマンド構築
    cmd = [
        'pyinstaller',
        '--onefile',
        '--clean',
        '--name=' + app_name,
        '--windowed'  # GUIアプリケーションなのでコンソールを表示しない
    ]
    
    # アイコンファイルが存在する場合は追加
    if icon_file:
        cmd.append('--icon=' + icon_file)
      # エントリポイントを追加
    cmd.append(entry_point)
    
    # ビルド実行
    print(f"実行コマンド: {' '.join(cmd)}")
    try:
        # capture_output=Falseでリアルタイムに出力を表示
        result = subprocess.run(cmd, check=True, capture_output=False)
        print(f"\nビルド成功！distフォルダに{app_name}{ext}が作成されました。")
    except subprocess.CalledProcessError as e:
        print(f"\nビルド中にエラーが発生しました: {e}")
        return 1
        
    return 0

if __name__ == "__main__":
    sys.exit(main())
