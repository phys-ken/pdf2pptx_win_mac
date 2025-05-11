# 開発者ノート

PDF to PPTX Converter開発プロジェクトへの参加を検討されている皆様へ。このドキュメントでは、開発プロセス、コントリビューションガイドライン、およびプロジェクトの進行状況について説明します。

## 開発環境のセットアップ

### 前提条件
- Python 3.8以上
- Git
- 開発用エディタ (Visual Studio Code推奨)

### リポジトリのクローンと初期設定

```bash
# リポジトリのクローン
git clone https://github.com/phys-ken/pdf2pptx_win_mac.git

# プロジェクトディレクトリに移動
cd pdf2pptx_win_mac

# 仮想環境を作成
python -m venv venv

# 仮想環境を有効化（Windowsの場合）
venv\Scripts\activate

# 仮想環境を有効化（macOS/Linuxの場合）
# source venv/bin/activate

# 依存パッケージのインストール
pip install -r requirements.txt
```

### 開発用コマンド

```bash
# GUIアプリケーションを実行
python src/pdf2pptx_gui.py

# コマンドラインから変換を実行
python src/pdf_converter.py <PDFファイルパス>

# テストを実行
python -m unittest discover -s src/tests

# 実行ファイルをビルド
python build.py
```

## プロジェクト構造

```
pdf2pptx_win_mac/
├── src/                 # ソースコード
│   ├── main/            # Electronメインプロセス
│   │   ├── index.js     # エントリーポイント
│   │   ├── menu.js      # アプリケーションメニュー
│   │   └── ipc.js       # IPC通信ハンドラー
│   │
│   ├── renderer/        # レンダラープロセス（UI）
│   │   ├── components/  # UIコンポーネント
│   │   ├── styles/      # CSS/SCSS
│   │   └── index.html   # メインHTML
│   │
│   ├── shared/          # 共有モジュール
│   │   ├── config.js    # 設定
│   │   └── utils.js     # ユーティリティ関数
│   │
│   └── modules/         # 機能モジュール
│       ├── pdf-parser/  # PDF解析モジュール
│       └── pptx-gen/    # PPTX生成モジュール
│
├── assets/              # 静的アセット
│   ├── icons/           # アプリケーションアイコン
│   └── images/          # その他の画像
│
├── docs/                # ドキュメント
│
├── scripts/             # ビルド/開発スクリプト
│
├── tests/               # テストコード
│
└── .config/             # 設定ファイル
    ├── webpack/         # Webpack設定
    └── electron/        # Electron設定
```

## コントリビューションガイドライン

### コントリビューションワークフロー

1. GitHubでリポジトリをフォーク
2. フィーチャーブランチを作成 (`git checkout -b feature/your-feature-name`)
3. コードの変更を実装
4. テストを記述して実行
5. 変更をコミットし、フォークに Push
6. オリジナルリポジトリに対してプルリクエストを作成

### コーディング規約

- **JavaScript**: [Airbnb JavaScript Style Guide](https://github.com/airbnb/javascript)に準拠
- **コミットメッセージ**: [Conventional Commits](https://www.conventionalcommits.org/)の形式に従う
- **ドキュメンテーション**: 新機能やAPI変更にはJSDocコメントとマークダウンドキュメントを更新

### プルリクエスト要件

プルリクエストは以下の基準を満たす必要があります：

1. コードは既存のテストをパスすること
2. 新機能には新しいテストが含まれていること
3. ドキュメントが更新されていること
4. ESLintとPrettierのチェックをパスすること
5. 変更内容の明確な説明があること

### 開発の流れ

1. GitHubのIssuesで課題や機能要求を確認
2. 取り組むIssueにコメントを残し、作業開始の意思を表明
3. 開発・テスト・ドキュメンテーション
4. プルリクエスト提出
5. コードレビューとフィードバックへの対応

## 現在の開発状況と優先事項

### アクティブな開発領域

現在、以下の領域での開発が活発に行われています：

1. **UI/UXの改善**
   - より直感的なユーザーインターフェースの開発
   - ダークモードのサポート

2. **変換品質の向上**
   - テキスト認識とエクスポート機能
   - 変換アルゴリズムの最適化

3. **パフォーマンス改善**
   - 大規模PDFのメモリ使用量最適化
   - 変換速度の向上

### 今後の計画

以下の機能は近い将来に実装予定です：

1. **バージョン1.5**（2023年第2四半期予定）
   - テキスト認識と抽出機能
   - 複数PDFの統合オプション
   - UIのリニューアル

2. **バージョン2.0**（2023年第4四半期予定）
   - PowerPointテンプレートの適用機能
   - クラウドストレージ統合
   - Linuxサポート

## 既知の課題とバグ

現在把握している主な課題は以下の通りです：

1. **大きなPDF（100MB以上）処理時のメモリ使用量が過大**
   - Issue #42で追跡中
   - メモリ使用量最適化の実装を検討中

2. **特定の複雑なPDFでレンダリングが不正確**
   - Issue #57で追跡中
   - pdf.jsのアップデートで部分的に対応予定

3. **macOS M1/M2チップでの一部最適化問題**
   - Issue #63で追跡中
   - ネイティブモジュールの最適化を検討中

## コミュニケーション

### コミュニティチャンネル

- **GitHub Discussions**: 一般的な議論と質問
- **GitHub Issues**: バグ報告と機能リクエスト
- **プロジェクトWiki**: 詳細なドキュメントと設計ガイド

### 定期ミーティング

コントリビューターとのミーティングは隔週水曜日に開催されています。参加を希望する場合は、GitHubのディスカッションで詳細をご確認ください。

## アプリケーションのビルド方法

PDF to PPTX Converterは、PyInstallerを使用してスタンドアロンの実行ファイルにビルドできます。ビルドプロセスは`build.py`スクリプトを使って自動化されています。

### ビルド前準備

1. 仮想環境をセットアップします：

```bash
# 仮想環境を作成
python -m venv venv

# Windows環境で仮想環境を有効化
.\venv\Scripts\Activate.ps1
# または、Command Promptの場合：
# .\venv\Scripts\activate.bat

# macOS/Linux環境で仮想環境を有効化
# source venv/bin/activate

# 必要なパッケージをインストール
pip install -r requirements.txt

# ビルドに必要なPyInstallerをインストール
pip install pyinstaller
```

### ビルド実行

セットアップが完了したら、以下のコマンドでビルドを実行します：

```bash
python build.py
```

ビルドが成功すると、`dist`フォルダに`pdf2pptx_converter.exe`（Windows）または`pdf2pptx_converter`（macOS/Linux）が作成されます。

### ビルド設定のカスタマイズ

デフォルトの設定をカスタマイズする場合は、`build.py`ファイルを編集します：

- アプリケーション名：`app_name`変数を変更
- アイコン：`resources`フォルダに`app_icon.ico`ファイルを配置（Windows）または`app_icon.icns`（macOS）
- その他のPyInstallerオプション：`cmd`リストに追加パラメータを追加

### ビルドのトラブルシューティング

ビルド中に問題が発生した場合：

1. PyInstallerのバージョンが最新であることを確認
2. 必要なすべての依存関係がインストールされていることを確認
3. `build`フォルダを削除して再試行
4. 詳細なデバッグ出力を得るには、`build.py`の`subprocess.run`呼び出しに`capture_output=False`を追加

---

このプロジェクトへの貢献に興味をお持ちいただき、ありがとうございます！質問や提案がある場合は、GitHub Discussionsまたは[Issue](https://github.com/phys-ken/pdf2pptx_win_mac/issues)でお気軽にご連絡ください。
