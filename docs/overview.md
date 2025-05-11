# PDF to PPTX Converter - プロジェクト概要

<p align="center">
  <img src="../resources/project_banner.png" alt="プロジェクトバナー" width="700">
</p>

## 目次

1. [プロジェクトについて](#プロジェクトについて)
2. [開発の背景と目的](#開発の背景と目的)
3. [技術仕様](#技術仕様)
4. [機能の詳細](#機能の詳細)
5. [プロジェクトの構成](#プロジェクトの構成)
6. [技術的な特徴](#技術的な特徴)
7. [将来の展望](#将来の展望)
8. [関連リソース](#関連リソース)

## プロジェクトについて

PDF to PPTX Converterは、PDFファイルをPowerPointプレゼンテーション(PPTX)に変換するためのオープンソースツールです。当初は教育関係者向けに開発されましたが、ビジネスシーンや研究発表など、様々な場面で活用できるよう拡張されています。

このプロジェクトの核となる機能は次の通りです：

- PDFファイルの各ページを高品質な画像に変換
- 生成した画像をPowerPointスライドとして配置
- 元のPDFと同じ寸法でスライドを作成
- シンプルで直感的なユーザーインターフェース

<p align="center">
  <img src="../resources/workflow_diagram.png" alt="ワークフロー図" width="600">
</p>

## 開発の背景と目的

### 課題の発見

教育現場や企業プレゼンテーションにおいて、以下のような課題がありました：

1. **既存のPDF資料の活用ニーズ**
   - 教科書や参考資料などのPDFをPowerPointで使いたい
   - 既存の報告書をプレゼン資料として再利用したい

2. **既存ソリューションの問題点**
   - 有料ソフトは予算的に導入が難しい
   - 無料ツールは機能が限定的または広告が多い
   - コマンドラインツールは一般ユーザーには扱いにくい
   - GUI付きツールでも操作が複雑で混乱を招く

### 開発目標

これらの課題を解決するため、以下の目標を掲げて開発を開始しました：

1. **シンプルさを極める**
   - 必要最低限の機能に絞り込み
   - 分かりやすいユーザーインターフェース

2. **高い互換性と安定性**
   - 様々なPDF形式に対応
   - クラッシュやエラーを最小限に

3. **すべての人が使える**
   - 無料・オープンソースで提供
   - 技術者でなくても直感的に操作可能

4. **軽量かつ高速**
   - インストール容量の最小化
   - 変換処理の最適化

## 技術仕様

### システム要件

- **対応OS**: Windows 10/11 (64-bit)
- **必要ディスク容量**: 約50MB
- **必要メモリ**: 最低2GB以上推奨

### 使用技術

当プロジェクトでは、以下の技術を採用しています：

| 要素 | 使用技術 | 用途 |
|------|---------|------|
| 開発言語 | Python 3.8+ | アプリケーション全体の開発 |
| PDF処理 | PyMuPDF (fitz) | PDFファイルの処理と画像変換 |
| PPTX生成 | python-pptx | PowerPointファイルの作成 |
| GUI | tkinter, tkinterdnd2 | グラフィカルインターフェース |
| 画像処理 | Pillow (PIL) | 画像サイズ調整と保存 |
| パッケージング | PyInstaller | スタンドアロン実行ファイルの作成 |

<p align="center">
  <img src="../resources/tech_stack.png" alt="技術スタック" width="500">
</p>

### パフォーマンス特性

- **変換速度**: 約1秒/ページ（標準的なテキストベースのPDF）
- **メモリ使用量**: ページ数と画像品質に依存（最大約300MB）
- **出力品質**: 元のPDFの解像度に依存（標準DPI: 300）

## 機能の詳細

### 主要機能

#### 1. PDF読み込み・変換機能

- 単一PDFファイルの読み込み
- PyMuPDFを使用した高品質な画像変換
- ページサイズの自動検出

#### 2. PowerPoint生成機能

- PDFのサイズに合わせたスライドサイズ設定
- 各ページを別々のスライドとして配置
- 画像サイズの自動最適化

#### 3. 画像保存機能

- PDFの各ページをJPEG画像として保存
- 後で再利用可能な形式での保存
- フォルダ構造の自動生成

#### 4. ユーザーインターフェース

- ドラッグ＆ドロップによるファイル選択
- 進捗状況のリアルタイム表示
- 直感的なボタン配置とシンプルな操作性

<p align="center">
  <img src="../resources/features_diagram.png" alt="機能概要図" width="600">
</p>

## プロジェクトの構成

プロジェクトは以下のような構成になっています：

```
pdf2pptx_win_mac/
├── src/                     # ソースコード
│   ├── pdf_converter.py     # PDF変換エンジン
│   ├── pdf2pptx_gui.py      # GUIアプリケーション
│   ├── folder_controller.py # プロジェクト管理
│   ├── hooks/               # PyInstallerフック
│   └── tests/               # テストコード
├── docs/                    # ドキュメント
│   ├── overview.md          # 本ドキュメント
│   └── manual.md            # ユーザーマニュアル
├── resources/               # 画像などのリソース
├── legacy/                  # 旧バージョンのコード
├── README.md                # プロジェクト説明
└── requirements.txt         # 依存パッケージ
```

詳細については、[GitHub リポジトリ](https://github.com/phys-ken/pdf2pptx_win_mac)をご覧ください。

## 技術的な特徴

### アーキテクチャ

このプロジェクトは、ビジネスロジック（PDF変換処理）とプレゼンテーション層（GUI）を明確に分離したモジュラー設計を採用しています。

<p align="center">
  <img src="../resources/architecture.png" alt="アプリケーションアーキテクチャ" width="550">
</p>

### コアロジック

```python
# PDF変換処理の核となる部分（簡略化）
def convert_pdf_to_images(pdf_path):
    pdf_document = fitz.open(pdf_path)
    images = []
    
    for page in pdf_document:
        pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    
    return images

def create_pptx(images, output_path):
    ppt = Presentation()
    
    # 最初の画像からスライドサイズを設定
    width, height = images[0].size
    ppt.slide_width = Pt(width)
    ppt.slide_height = Pt(height)
    
    for img in images:
        slide = ppt.slides.add_slide(ppt.slide_layouts[6])
        slide.shapes.add_picture(img, 0, 0, width=ppt.slide_width, height=ppt.slide_height)
    
    ppt.save(output_path)
```

### 工夫点

1. **メモリ効率の最適化**
   - 各ページを順次処理し、メモリ消費を抑制
   - 不要なリソースの積極的な解放

2. **エラーハンドリング**
   - 様々なPDF形式への対応
   - 破損ファイルや非標準PDFの検出と適切なエラーメッセージ

3. **UIの応答性維持**
   - 変換処理は別スレッドで実行
   - プログレスバーによるリアルタイム進捗表示

4. **サイズ制限への対応**
   - PowerPointのスライドサイズ制限（1-56インチ）に対応
   - 自動的なスケーリング処理による最適化

## 将来の展望

### 短期的な改善計画

- **バッチ処理機能** - 複数PDFの一括変換
- **設定オプション** - 解像度や出力形式のカスタマイズ
- **多言語対応** - 英語、中国語などのローカライズ

### 中長期的なロードマップ

1. **オンラインサービス化**
   - Webブラウザからの変換に対応
   - クラウド上でのバッチ処理

2. **機能拡張**
   - PDFテキストの抽出と編集可能なスライド作成
   - OCRを活用したテキストレイヤー追加
   - カスタムテンプレート対応

3. **クロスプラットフォーム対応**
   - Mac版のリリース
   - Linuxサポート
   - モバイル対応（Android/iOS）

## 関連リソース

- [ユーザーマニュアル](./manual.md) - 詳細な使用方法について
- [ライセンス情報](../LICENSE) - MITライセンスについて
- [コントリビューションガイド](https://github.com/phys-ken/pdf2pptx_win_mac/blob/main/CONTRIBUTING.md) - 開発への参加方法
- [リリースノート](https://github.com/phys-ken/pdf2pptx_win_mac/releases) - 各バージョンの変更点

---

作成者: phys-ken ([@phys_ken](https://twitter.com/phys_ken))  
最終更新: 2025年5月11日
