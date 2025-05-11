"""
フォルダ構造を管理するためのモジュール
重複Markdownファイルを整理し、必要最小限の構成に整理します
"""
import os
import shutil
import sys
import glob


def create_folders():
    """必要なフォルダ構造を作成する"""
    folders = [
        'legacy',
        'resources',
        'src',
        'src/tests',
        'docs',  # ドキュメントフォルダ
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
    src_dir = os.path.join(root_dir, 'src')
    
    folders_to_remove = [
        os.path.join(root_dir, 'dist'),
        os.path.join(root_dir, 'dist_win'),
        os.path.join(root_dir, 'dist_mac'),
        os.path.join(root_dir, 'build'),
        os.path.join(root_dir, '__pycache__'),
        os.path.join(root_dir, 'output'),
        os.path.join(src_dir, 'figtmpfig'),
        os.path.join(src_dir, 'hooks'),
    ]
    
    for folder in folders_to_remove:
        if os.path.exists(folder):
            try:
                shutil.rmtree(folder)
                print(f"フォルダを削除: {folder}")
            except Exception as e:
                print(f"フォルダの削除に失敗: {folder} - {e}")


def organize_markdown_files():
    """Markdownファイルを整理し、重複を解消する"""
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    docs_dir = os.path.join(root_dir, 'docs')
    
    # 1. すべてのMarkdownファイルを収集 (READMEファイル以外)
    md_files = []
    for root, _, files in os.walk(root_dir):
        for file in files:
            if file.lower().endswith('.md') and file.lower() != 'readme.md':
                md_files.append(os.path.join(root, file))
    
    print(f"検出されたMarkdownファイル (README除く): {len(md_files)}個")
    for md_file in md_files:
        print(f" - {md_file}")
    
    # 3. docsフォルダ内のファイルを整理
    docs_md_files = glob.glob(os.path.join(docs_dir, '*.md'))
    
    # 3.1. 既存のユーザーガイド関連ファイルを検出
    user_guide_files = [f for f in docs_md_files if any(x in os.path.basename(f).lower() 
                                                      for x in ['user', 'guide', 'manual', 'usage'])]
    
    # 3.2. 技術情報関連ファイルを検出
    tech_files = [f for f in docs_md_files if any(x in os.path.basename(f).lower() 
                                                for x in ['tech', 'detail', 'overview'])]
    
    # 3.3. FAQ関連ファイルを検出
    faq_files = [f for f in docs_md_files if 'faq' in os.path.basename(f).lower()]
    
    # 4. ファイルを統合する
    if user_guide_files:
        create_unified_user_guide(user_guide_files, docs_dir)
    
    if tech_files:
        create_unified_tech_doc(tech_files, docs_dir)
    
    if faq_files:
        create_unified_faq(faq_files, docs_dir)


def create_unified_user_guide(source_files, docs_dir):
    """ユーザーガイドを統合する"""
    output_path = os.path.join(docs_dir, 'user-guide.md')
    
    content = """# PDF to PPTX Converter - ユーザーガイド

## 目次
1. インストール方法
2. 基本的な使い方
3. オプション機能
4. よくある問題と解決方法

"""
    
    # 既存のファイルから内容を抽出して統合
    for file_path in source_files:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                file_content = f.read()
                content += f"## {os.path.basename(file_path)} からの内容\n\n"
                content += file_content + "\n\n"
        except Exception as e:
            print(f"ファイル読み込みエラー: {file_path} - {e}")
    
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"統合されたユーザーガイドを作成: {output_path}")
        
        # 元のファイルを削除
        for file_path in source_files:
            if os.path.exists(file_path) and file_path != output_path:
                os.remove(file_path)
                print(f"統合後、元ファイルを削除: {file_path}")
    except Exception as e:
        print(f"ユーザーガイド作成エラー: {e}")


def create_unified_tech_doc(source_files, docs_dir):
    """技術文書を統合する"""
    output_path = os.path.join(docs_dir, 'technical-details.md')
    
    content = """# PDF to PPTX Converter - 技術情報

## 目次
1. システム概要
2. アーキテクチャ
3. 使用ライブラリ
4. 実装の詳細

"""
    
    # 既存のファイルから内容を抽出して統合
    for file_path in source_files:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                file_content = f.read()
                content += f"## {os.path.basename(file_path)} からの内容\n\n"
                content += file_content + "\n\n"
        except Exception as e:
            print(f"ファイル読み込みエラー: {file_path} - {e}")
    
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"統合された技術文書を作成: {output_path}")
        
        # 元のファイルを削除
        for file_path in source_files:
            if os.path.exists(file_path) and file_path != output_path:
                os.remove(file_path)
                print(f"統合後、元ファイルを削除: {file_path}")
    except Exception as e:
        print(f"技術文書作成エラー: {e}")


def create_unified_faq(source_files, docs_dir):
    """FAQを統合する"""
    output_path = os.path.join(docs_dir, 'faq.md')
    
    content = """# PDF to PPTX Converter - よくある質問

"""
    
    # 既存のファイルから内容を抽出して統合
    for file_path in source_files:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                file_content = f.read()
                content += file_content + "\n\n"
        except Exception as e:
            print(f"ファイル読み込みエラー: {file_path} - {e}")
    
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"統合されたFAQを作成: {output_path}")
        
        # 元のファイルを削除
        for file_path in source_files:
            if os.path.exists(file_path) and file_path != output_path:
                os.remove(file_path)
                print(f"統合後、元ファイルを削除: {file_path}")
    except Exception as e:
        print(f"FAQ作成エラー: {e}")


def update_readme():
    """READMEを更新しつつ、必ず保持する（削除しない）"""
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    old_readme_path = os.path.join(root_dir, 'readme.md')
    new_readme_path = os.path.join(root_dir, 'README.md')
    
    # まず現在のREADMEをバックアップ (READMEがどちらも存在しない場合に備えて)
    backup_created = False
    
    if os.path.exists(old_readme_path):
        readme_backup_path = os.path.join(root_dir, 'readme.md.bak')
        try:
            shutil.copy2(old_readme_path, readme_backup_path)
            backup_created = True
            print(f"readme.mdのバックアップを作成: {readme_backup_path}")
        except Exception as e:
            print(f"readme.mdのバックアップ作成失敗: {e}")
    
    if os.path.exists(new_readme_path):
        readme_backup_path = os.path.join(root_dir, 'README.md.bak')
        try:
            shutil.copy2(new_readme_path, readme_backup_path)
            backup_created = True
            print(f"README.mdのバックアップを作成: {readme_backup_path}")
        except Exception as e:
            print(f"README.mdのバックアップ作成失敗: {e}")
    
    # 古いreadme.mdが存在し、新しいREADME.mdが存在しない場合は名前を変更
    if os.path.exists(old_readme_path) and not os.path.exists(new_readme_path):
        try:
            shutil.copy2(old_readme_path, new_readme_path)
            print(f"readme.md から README.md を作成")
        except Exception as e:
            print(f"README.md作成失敗: {e}")
    
    # 両方存在する場合は統合または選択
    elif os.path.exists(old_readme_path) and os.path.exists(new_readme_path):
        try:
            # ファイルサイズを比較して、より大きい/詳細なファイルを残す
            old_size = os.path.getsize(old_readme_path)
            new_size = os.path.getsize(new_readme_path)
            
            if old_size > new_size:
                # 古いファイルの方が大きければそれを使用
                os.remove(new_readme_path)
                shutil.copy2(old_readme_path, new_readme_path)
                print("より詳細なreadme.mdをREADME.mdとして使用")
            else:
                # 新しいファイルの方が大きいか同じなら、readme.mdをバックアップ
                print("既存のREADME.mdを維持し、readme.mdをバックアップ")
        except Exception as e:
            print(f"READMEファイル処理エラー: {e}")
    
    # バックアップから復元が必要な場合
    if not os.path.exists(new_readme_path) and backup_created:
        try:
            # バックアップファイルパスを検索
            backup_files = glob.glob(os.path.join(root_dir, '*.md.bak'))
            if backup_files:
                latest_backup = max(backup_files, key=os.path.getmtime)
                shutil.copy2(latest_backup, new_readme_path)
                print(f"バックアップからREADME.mdを復元: {latest_backup}")
        except Exception as e:
            print(f"バックアップからの復元に失敗: {e}")
    
    # どうしてもREADME.mdがない場合は新規作成
    if not os.path.exists(new_readme_path):
        create_default_readme(new_readme_path)


def create_default_readme(readme_path):
    """デフォルトのREADME.mdを作成する"""
    content = """# PDF to PPTX Converter

PDFファイルをPowerPointプレゼンテーションに変換するシンプルなツール。各ページを1枚のスライドに配置します。

## 特徴

- シンプルなGUIインターフェース
- ドラッグ＆ドロップでPDFファイルを簡単に変換
- PowerPointスライドサイズを元のPDFサイズに自動調整
- 出力先フォルダの選択が可能（デフォルトはPDFと同じ場所）

## インストール

[リリースページ](https://github.com/phys-ken/pdf2pptx_win_mac/releases)から最新版をダウンロードして実行するだけです。

## 詳細情報

詳細な使い方やプロジェクトの概要については、[ドキュメント](docs/technical-details.md)をご覧ください。

## ライセンス

MITライセンスの下で公開されています。
"""
    
    try:
        with open(readme_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"新しいREADME.mdを作成しました: {readme_path}")
    except Exception as e:
        print(f"README.md作成エラー: {e}")


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


def main():
    """メイン実行関数"""
    print("フォルダ構成の整理を開始します...")
    
    # 実行パスの取得と修正
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    global src_dir
    src_dir = os.path.join(root_dir, 'src')
    
    # 1. READMEファイルの処理を最初に行う（バックアップを作成）
    update_readme()
    
    # 2. 基本的なフォルダ構造の作成
    create_folders()
    
    # 3. レガシーファイルの移動
    move_legacy_files()
    
    # 4. 不要なフォルダを削除
    cleanup_folders()
    
    # 5. Markdownファイルの整理（重複解消）
    organize_markdown_files()
    
    # 6. 必要なファイルの確認
    check_necessary_files()
    
    # 7. READMEが正しく存在するか最終確認
    if not os.path.exists(os.path.join(root_dir, 'README.md')):
        print("警告: README.mdが見つかりません。再作成します。")
        create_default_readme(os.path.join(root_dir, 'README.md'))
    
    print("フォルダ構成の整理が完了しました")


if __name__ == "__main__":
    main()