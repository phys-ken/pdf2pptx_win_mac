# PDFをPPTXに一枚ずつ貼り付けます。
## できること
* pdfをjpegに変換します(内部処理)
* pptxファイルを作成します。スライドの縦横サイズを、PDFの一枚目と合わせます
* output.pptxを作成します。

## 使い方
* osに合わせて、dist_macかdist_winの中にあるファイルをダウンロードする。
* ダブルクリックして起動(長いと2分くらいかかる。)
* コマンドプロンプトに入力欄が表示されたら、パスを入力する。

## 技術的なリファレンス
* windowsにpopplerをインストールする。
  * [stack overflow How to install Poppler on Windows?](https://stackoverflow.com/questions/18381713/how-to-install-poppler-on-windows)
    * 2020年6月27日の　Owen Schwartz　のアンサーが参考になった。popplerの最新版をbuildして、[Github](https://github.com/oschwartz10612/poppler-windows/releases)に上げてくれている。
    * zipを展開して、C:user~ programfileの中にいれて、環境変数のPATHを通す。

* pdf2pptxをインストールする。
  * popplerの環境変数を通した後に、`pip install pdf2image` でインストール。

* pyinstallerで、.pyを実行可能ファイルに変える。
  * [teratail Pythonで、pipしたパッケージをインポートしている.pyファイルをexe化しても実行できない？](https://teratail.com/questions/184343)を参考に、hook-pptx.pyを作成
  * 以下のコードを実行

```
pyinstaller --onefile --additional-hooks-dir hooks  pdf2pptx.py       
```

* 実行可能ファイルが重すぎるとき...
  * numpyやpandasも含まれてしまうらしい。仮想環境を変えて、必要なモジュールのみpip installしなおす。
  * [この通り](https://qiita.com/napinoco/items/068ce8ef6ef4309966b1)にしました。