# PDF to PPTX Converter - 技術情報

## 目次
1. システム概要
2. アーキテクチャ
3. 使用ライブラリ
4. 実装の詳細

## overview.md からの内容

# PDF to PPTX Converter プロジェクト概要

## プロジェクトについて

PDF to PPTX Converterは、PDFファイルをMicrosoft PowerPointのプレゼンテーションに簡単に変換するためのオープンソースツールです。教育者、ビジネスプロフェッショナル、研究者など、PDFコンテンツをプレゼンテーション形式で活用したいユーザーのために開発されました。

## 目次

- [プロジェクトの目的](#プロジェクトの目的)
- [開発の経緯](#開発の経緯)
- [主な機能](#主な機能)
- [技術スタック](#技術スタック)
- [プロジェクト構成](#プロジェクト構成)
- [今後の展望](#今後の展望)
- [関連ドキュメント](#関連ドキュメント)

## プロジェクトの目的

PDF to PPTX Converterは以下の目的で開発されました：

1. **アクセシビリティの向上**: PDFコンテンツをPowerPoint形式で編集可能にすることで、教材や資料のカスタマイズを容易にする
2. **ワークフロー効率化**: 既存のPDFデータを活用したプレゼンテーション作成を効率化する
3. **クロスプラットフォーム対応**: WindowsとmacOSの両環境で一貫した変換体験を提供する
4. **シンプルさの追求**: 技術的な知識がなくても簡単に使用できるインターフェースを提供する

## 開発の経緯

このプロジェクトは、教育現場での実際のニーズから生まれました。多くの教師や講師が、論文やウェブ上のPDF資料をプレゼンテーションに組み込む際の煩雑さに直面していました。既存のソリューションは複雑で高価、または変換品質が低いという課題がありました。

初期バージョンは2021年に内部ツールとして開発され、教育機関での使用を通じてフィードバックを収集し、改良を重ねてきました。2022年にオープンソースプロジェクトとして公開され、コミュニティからの貢献を受けて機能拡張とバグ修正が継続的に行われています。

## 主な機能

- **直感的なユーザーインターフェース**: ドラッグ＆ドロップによる簡単なファイル選択
- **高品質な変換**: PDFの視覚的要素を忠実に再現
- **柔軟な出力オプション**: スライドサイズや画質の調整が可能
- **バッチ処理**: 複数のPDFファイルを一括変換
- **クロスプラットフォーム対応**: Windows, macOS両環境で同様の機能を提供

## 技術スタック

PDF to PPTX Converterは以下の技術を使用して開発されています：

- **フロントエンド**: Electron.js（クロスプラットフォームGUIアプリケーション）
- **PDF処理**: pdf.js, pdf-lib（PDFの解析と処理）
- **PPTX生成**: pptxgenjs（PowerPointファイルの生成）
- **画像処理**: Sharp（高速な画像変換と最適化）
- **ビルド/パッケージング**: Electron-builder（配布可能なアプリケーションの作成）

## プロジェクト構成

プロジェクトは以下のような構成になっています：

```
pdf2pptx_win_mac/
├── src/                # ソースコード
│   ├── main/           # メインプロセス（Electron）
│   ├── renderer/       # レンダラープロセス（UI関連）
│   └── shared/         # 共有モジュール
├── assets/             # アイコンなどの静的アセット
├── docs/               # ドキュメント
├── scripts/            # ビルドスクリプトなど
└── tests/              # テストコード
```

## 今後の展望

PDF to PPTX Converterの今後の開発ロードマップには以下の機能が含まれています：

1. **テキスト認識と編集可能なテキストの生成**: OCR技術を統合し、PDFのテキストを編集可能な形式で抽出
2. **複数PDFの結合**: 複数のPDFファイルを1つのPPTXプレゼンテーションに統合
3. **スタイルテンプレート**: カスタムテンプレートを適用したPPTX生成
4. **クラウド統合**: クラウドストレージサービスとの連携
5. **Linuxサポート**: Ubuntu, Fedoraなどの主要Linuxディストリビューションへの対応

## 関連ドキュメント

- [ユーザーガイド](user-guide.md) - 詳細な使用方法
- [よくある質問 (FAQ)](faq.md) - 一般的な質問と回答
- [技術情報](technical-details.md) - 内部実装の詳細
- [開発者ノート](developer-notes.md) - 開発参加のガイドライン

---

このプロジェクトに貢献してくださる方を歓迎します。バグレポートや機能リクエストは[GitHubのIssuesページ](https://github.com/phys-ken/pdf2pptx_win_mac/issues)までお願いします。


## technical-details.md からの内容

# PDF to PPTX Converter - 技術情報

## 目次
1. システム概要
2. アーキテクチャ
3. 使用ライブラリ
4. 実装の詳細

## システム概要

PDF to PPTX Converterは、PyMuPDF (fitz) ライブラリを使用してPDFファイルをページ単位で画像に変換し、python-pptxを使用してPowerPointプレゼンテーションを作成するPythonアプリケーションです。GUIはtkinterとtkinterdnd2を使用して実装されています。

## アーキテクチャ

アプリケーションは以下のコンポーネントで構成されています：

### コンポーネント構成

1. **GUI層** (`pdf2pptx_gui.py`)
   - ユーザーインターフェース
   - ドラッグ＆ドロップ機能
   - ファイル選択ダイアログ
   - 進捗表示

2. **PDF変換エンジン** (`pdf_converter.py`)
   - PDFの読み込み (PyMuPDF)
   - ページの画像変換
   - PowerPointファイル生成 (python-pptx)

3. **プロジェクト管理** (`folder_controller.py`)
   - フォルダ構造の管理
   - ファイルの整理

### 処理フロー

1. PDFファイルを読み込む
2. 各ページを画像 (JPEG) に変換
3. PowerPointのスライドサイズを元のPDF寸法に合わせて設定
4. 各画像をスライドに貼り付け
5. PowerPointファイルとして保存

## 使用ライブラリ

- **PyMuPDF (fitz)**: PDFファイルの読み取りと画像変換
- **python-pptx**: PowerPointファイルの作成と操作
- **tkinter**: GUIの基本フレームワーク
- **tkinterdnd2**: ドラッグ＆ドロップ機能
- **Pillow (PIL)**: 画像処理
- **threading**: 非同期処理による応答性の向上
- **os, sys, shutil**: ファイルシステム操作

## 実装の詳細

### PDF処理

PDFの処理には、PyMuPDF (fitz) ライブラリを使用しています。これにより、以下の操作が可能です：

```python
# PDFドキュメントの読み込み
doc = fitz.open(pdf_path)

# ページをJPEG画像に変換
for page_num in range(len(doc)):
    page = doc.load_page(page_num)
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    pix.save(image_path, "jpeg")
```

### PowerPoint生成

PowerPointファイルの生成には、python-pptxライブラリを使用しています：

```python
# PowerPointプレゼンテーションの作成
prs = Presentation()
slide_width = pdf_width * POINTS_PER_INCH
slide_height = pdf_height * POINTS_PER_INCH
prs.slide_width = slide_width
prs.slide_height = slide_height

# スライドの追加と画像の配置
for image_file in image_files:
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 白紙レイアウト
    slide.shapes.add_picture(image_file, 0, 0, width=slide_width, height=slide_height)
```

### パフォーマンス最適化

- 大きなPDFファイルを効率的に処理するためのチャンク処理
- 変換中のメモリ管理の最適化
- マルチスレッドによる応答性の向上

### エラー処理

アプリケーションには様々なエラーを処理する仕組みが組み込まれています：

- ファイルアクセスエラー
- メモリ不足エラー
- 変換エラー
- プロセス中断

エラーが発生した場合、ユーザーに適切なエラーメッセージを表示し、可能な場合は解決策を提示します。

## ビルド方法

PyInstallerを使用して、実行可能ファイルにビルドします：

```bash
# Windows用ビルド
pyinstaller --noconsole --add-data "resources/*;resources" --icon=resources/icon.ico pdf2pptx_gui.py

# macOS用ビルド
pyinstaller --noconsole --add-data "resources/*:resources" --icon=resources/icon.icns pdf2pptx_gui.py
```

## 今後の拡張計画

1. 変換設定のカスタマイズ機能 (解像度、画質など)
2. バッチ処理の強化
3. OCR機能の統合による検索可能なテキスト抽出
4. ダークモード対応


