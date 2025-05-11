# PDF to PPTX Converter

PDFファイルをPowerPoint (PPTX) ファイルに変換するデスクトップアプリケーションです。

このアプリケーションは、PDFの各ページを画像として抽出し、それらをPowerPointの各スライドに貼り付けます。
プログラミングの知識がない方でも簡単に使用できるように、シンプルなGUIを提供しています。

## 主な機能

- PDFファイルをドラッグ＆ドロップまたはファイル選択ダイアログで指定
- 各PDFページを画像（JPGまたはPNG）に変換
- 変換された画像をPowerPointの各スライドに自動配置
- 出力先のフォルダを指定可能（指定しない場合はPDFと同じフォルダに出力）
- 変換の進捗状況を表示

## 動作環境

- Windows

## インストールと実行

実行ファイルは、リリースページからダウンロードできます。（まだリリースページはありません）
ダウンロードした実行ファイルをダブルクリックするだけでアプリケーションが起動します。

## 使用方法

アプリケーションの詳しい使い方は、[ユーザーマニュアル](./docs/user_manual.md) を参照してください。

## 技術的な詳細

アプリケーションの技術的な詳細については、[技術詳細](./docs/technical_details.md) を参照してください。

## 依存ライブラリ

このアプリケーションは、以下の主要なPythonライブラリを使用しています。

- Tkinter (tkinterdnd2): GUIの構築
- PyMuPDF (fitz): PDFの解析と画像への変換
- python-pptx: PowerPointファイルの生成
- Pillow: 画像処理

詳細な依存関係は `requirements.txt` を参照してください。

## ビルド方法

ソースコードから実行ファイルをビルドする方法は、`build.py` を参照してください。
PyInstaller を使用してビルドします。

```shell
python build.py
```

## ライセンス

このプロジェクトは [MIT License](./LICENSE) の下で公開されています。
また、一部のコンポーネントは追加のライセンス条件が適用される場合があります。詳細は [LICENSE-ADDITIONAL-INFO.md](./LICENSE-ADDITIONAL-INFO.md) および [THIRDPARTY-LICENSES.md](./THIRDPARTY-LICENSES.md) を確認してください。

## 貢献

バグ報告や機能リクエストは、GitHubのIssuesにお願いします。

