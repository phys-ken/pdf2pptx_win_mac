"""
フォルダ構造を管理するためのモジュール
"""
import os
import shutil
import sys


def create_folders():
    """必要なフォルダ構造を作成する"""
    folders = [
        'legacy',
        'resources',
        'src',
        'src/tests',
        'src/hooks',
        'docs',  # ドキュメントフォルダを追加
    ]

    # プロジェクトのルートディレクトリを取得
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    for folder in folders:
        folder_path = os.path.join(root_dir, folder)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            print(f"フォルダを作成: {folder_path}")


def move_legacy_files():
    """既存のファイルをlegacyフォルダに移動する"""
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    legacy_dir = os.path.join(root_dir, 'legacy')
    src_dir = os.path.join(root_dir, 'src')
    
    # 移動対象のファイル
    legacy_files = [
        os.path.join(src_dir, 'pdf2pptx_win.py'),
        os.path.join(src_dir, 'pdf2pptx.py'),
        os.path.join(src_dir, 'pdf2pptx_win.spec'),
        os.path.join(src_dir, 'pdf2pptx.spec'),
        # hooks/ フォルダ内のファイルはそのまま残しておく（必要な場合があるため）
    ]
    
    for file_path in legacy_files:
        if os.path.exists(file_path):
            filename = os.path.basename(file_path)
            destination = os.path.join(legacy_dir, filename)
            shutil.copy2(file_path, destination)
            print(f"ファイルをlegacyにコピー: {filename}")
            # コピー後に元のファイルを削除
            os.remove(file_path)
            print(f"元のファイルを削除: {file_path}")


def cleanup_folders():
    """不要なフォルダを削除する"""
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    folders_to_remove = [
        os.path.join(root_dir, 'dist'),
        os.path.join(root_dir, 'dist_win'),
        os.path.join(root_dir, 'dist_mac'),
        os.path.join(root_dir, 'build'),
        os.path.join(root_dir, '__pycache__'),
        os.path.join(root_dir, 'output'),  # outputフォルダを削除
        os.path.join(src_dir, 'figtmpfig'),  # 一時的な画像フォルダを削除
        os.path.join(src_dir, 'hooks'),  # hooksフォルダを削除
    ]
    
    for folder in folders_to_remove:
        if os.path.exists(folder):
            try:
                shutil.rmtree(folder)
                print(f"フォルダを削除: {folder}")
            except Exception as e:
                print(f"フォルダの削除に失敗: {folder} - {e}")


def remove_old_readme():
    """古いreadme.mdを削除して新しいREADME.mdを作成する"""
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    old_readme_path = os.path.join(root_dir, 'readme.md')
    
    # 古いreadme.mdを削除
    if os.path.exists(old_readme_path):
        try:
            os.remove(old_readme_path)
            print(f"ファイルを削除: {old_readme_path}")
        except Exception as e:
            print(f"ファイルの削除に失敗: {old_readme_path} - {e}")
    
    # 新しいREADME.mdを作成
    new_readme_path = os.path.join(root_dir, 'README.md')
    readme_content = """# PDF to PPTX Converter

PDFファイルをPowerPointプレゼンテーションに変換するシンプルなツール。各ページを1枚のスライドに配置します。

## 特徴

- シンプルなGUIインターフェース
- ドラッグ＆ドロップでPDFファイルを簡単に変換
- PowerPointスライドサイズを元のPDFサイズに自動調整
- 出力先フォルダの選択が可能（デフォルトはPDFと同じ場所）

## インストール

[リリースページ](https://github.com/phys-ken/pdf2pptx_win_mac/releases)から最新版をダウンロードして実行するだけです。

## 詳細情報

詳細な使い方やプロジェクトの概要については、[ドキュメント](docs/overview.md)をご覧ください。

## ライセンス

MITライセンスの下で公開されています。
"""
    
    with open(new_readme_path, 'w', encoding='utf-8') as f:
        f.write(readme_content)
    
    print(f"README.md を作成: {new_readme_path}")


def create_docs():
    """docsフォルダにドキュメントファイルを作成する"""
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    docs_dir = os.path.join(root_dir, 'docs')
    
    # 概要ドキュメントの作成
    overview_path = os.path.join(docs_dir, 'overview.md')
    overview_content = """# PDF to PPTX Converter - 概要

## プロジェクトについて

このプロジェクトは、PDFファイルを簡単にPowerPointプレゼンテーション(PPTX)に変換するためのツールです。元々はpopplerに依存したCLIツールでしたが、現在はPyMuPDFを使用し、使いやすいGUIを実装しています。

## 技術仕様

- **使用言語**: Python 3.8以上
- **主要ライブラリ**:
  - PyMuPDF (fitz): PDF処理
  - python-pptx: PowerPointファイル生成
  - tkinter, tkinterdnd2: GUIインターフェース
  - Pillow: 画像処理

## 開発背景

教育現場や業務でPDFをPowerPointに変換する必要性が頻繁にあることから開発されました。既存のソリューションは複雑であったり、有料であったりしたため、シンプルで無料のツールを目指しました。

## 機能の詳細

- PDFの各ページをJPEG画像に変換
- 画像をPowerPointスライドとして配置
- スライドサイズをPDFのサイズに合わせて自動調整
- 画像ファイルも別途保存可能

## プロジェクト構成

- **src/**: ソースコード
  - pdf_converter.py: PDF変換エンジン
  - pdf2pptx_gui.py: GUIアプリケーション
  - folder_controller.py: プロジェクト管理スクリプト
- **docs/**: ドキュメント
- **legacy/**: 以前のバージョンのコード
- **resources/**: アイコンなどのリソース

## 今後の展望

- バッチ処理機能の追加
- 変換オプションの拡充（解像度調整など）
- ダークモード対応
"""
    
    with open(overview_path, 'w', encoding='utf-8') as f:
        f.write(overview_content)
    
    print(f"overview.md を作成: {overview_path}")
    
    # 操作マニュアルの作成
    manual_path = os.path.join(docs_dir, 'manual.md')
    manual_content = """# PDF to PPTX Converter - 操作マニュアル

## インストール方法

1. [リリースページ](https://github.com/phys-ken/pdf2pptx_win_mac/releases)から最新版の実行ファイルをダウンロード
2. ダウンロードしたZIPファイルを任意の場所に展開
3. `pdf2pptx_gui.exe` をダブルクリックして起動

## 基本的な使い方

### PDFファイルの変換

1. アプリケーションを起動する
2. 以下のいずれかの方法でPDFファイルを選択:
   - PDFファイルをアプリケーションウィンドウにドラッグ＆ドロップ
   - 「PDFファイルを選択」ボタンをクリックしてファイル選択ダイアログから選択
3. 「変換開始」ボタンをクリック
4. 変換が完了すると、成功メッセージが表示される

### 出力先の指定

デフォルトでは、PDFファイルと同じフォルダに出力されます。
出力先を変更したい場合:

1. 「出力先を選択 (オプション)」ボタンをクリック
2. フォルダ選択ダイアログで保存先を指定

## 出力ファイル

変換が完了すると、以下のファイルが生成されます:

- `[PDFファイル名].pptx` - 変換されたPowerPointファイル
- `[PDFファイル名]_figs` - PDFの各ページの画像が保存されたフォルダ

## よくあるトラブルシューティング

### 実行時にエラーが表示される場合

- .NET Frameworkが最新でない可能性があります。Windows Updateを実行してください。
- セキュリティソフトがアプリケーションの実行をブロックしている可能性があります。例外設定を確認してください。

### PDFが正しく変換されない

- 非常に大きなサイズのPDFは変換に時間がかかる場合があります。しばらくお待ちください。
- パスワード保護されたPDFは変換できません。PDFのセキュリティ設定を解除してから変換してください。
- PDFの内容によっては、画像品質に影響が出ることがあります。

### その他の問題

その他の問題が発生した場合は、[GitHubのIssue](https://github.com/phys-ken/pdf2pptx_win_mac/issues)でご報告ください。
"""
    
    with open(manual_path, 'w', encoding='utf-8') as f:
        f.write(manual_content)
    
    print(f"manual.md を作成: {manual_path}")


def create_legacy_md():
    """legacyフォルダにlegacy.mdファイルを作成する"""
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    legacy_md_path = os.path.join(root_dir, 'legacy', 'legacy.md')
    
    content = """# 旧バージョンのPDF2PPTXについて

## 変更点
- GUIインターフェースの追加
- Poppler依存からPyMuPDFへの変更
- フォルダ構成の最適化
- クリックによるファイル選択の操作性向上

## 旧バージョンの課題
1. コマンドラインでの操作が必要（ユーザーフレンドリーでない）
2. Popplerへの依存があり、インストールが複雑
3. Windows版とMac版で別々の実装が必要
4. 出力先の指定が毎回必要

## 解決方法
- PyMuPDFを使用してPDF処理を行うことでPopplerへの依存をなくす
- tkinterによるシンプルなGUIの実装
- 出力先のデフォルト設定（PDFと同じフォルダ）
"""
    
    with open(legacy_md_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"legacy.md を作成: {legacy_md_path}")


def check_necessary_files():
    """必要なファイルが存在するか確認"""
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    src_dir = os.path.join(root_dir, 'src')
    
    necessary_files = [
        os.path.join(src_dir, 'pdf_converter.py'),
        os.path.join(src_dir, 'pdf2pptx_gui.py'),
    ]
    
    all_exist = True
    for file_path in necessary_files:
        if not os.path.exists(file_path):
            print(f"警告: 必要なファイルが見つかりません: {file_path}")
            all_exist = False
    
    if all_exist:
        print("必要なファイルはすべて存在しています。")


def create_hooks_file():
    """hooks/hook-pptx.pyファイルを作成する"""
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    hooks_dir = os.path.join(root_dir, 'src', 'hooks')
    
    # hooksディレクトリを作成（すでに作成済みの場合もあるが念のため）
    if not os.path.exists(hooks_dir):
        os.makedirs(hooks_dir)
        print(f"フォルダを作成: {hooks_dir}")
    
    # hook-pptx.pyファイルのパス
    hook_file_path = os.path.join(hooks_dir, 'hook-pptx.py')
    
    # hook-pptx.pyの内容
    hook_content = """# hook-pptx.py
# PyInstallerのhooks
from PyInstaller.utils.hooks import collect_data_files

# python-pptxに必要なデータファイルを収集
datas = collect_data_files('pptx')
"""
    
    # ファイルを書き込み
    with open(hook_file_path, 'w', encoding='utf-8') as f:
        f.write(hook_content)
    
    print(f"hook-pptx.py を作成: {hook_file_path}")


def main():
    """メイン実行関数"""
    print("フォルダ構成の整理を開始します...")
    # 実行パスの取得と修正（src_dirをグローバル変数として定義）
    global src_dir
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    src_dir = os.path.join(root_dir, 'src')
    
    create_folders()
    move_legacy_files()
    cleanup_folders()
    remove_old_readme()  # readme.mdの削除とREADME.mdの作成
    create_docs()        # docsフォルダとドキュメントの作成
    create_legacy_md()
    create_hooks_file()
    check_necessary_files()
    print("フォルダ構成の整理が完了しました")


if __name__ == "__main__":
    main()