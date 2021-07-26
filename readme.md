# PDFをPPTXに一枚ずつ貼り付けます。
## できること
* pdfをjpegに変換します(内部処理)
* pptxファイルを作成します。スライドの縦横サイズを、PDFの一枚目と合わせます
* output.pptxを作成します。

## 使い方
* [release](https://github.com/phys-ken/pptx2pdf_win_mac/releases)から、最新版をダウンロードしてください。
* 実行可能ファイルをダブルクリックして起動(長いと2分くらいかかる。)
* 起動が完了すると、コマンドプロンプトに`PDFファイルのフルパスを入力してください>>>`と表示されます。以降は、支持の通りに、入力・コピペをしてください。
  * **ドラッグ＆ドロップ**でもOK!

## 開発時のメモ
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

## popplerをpyinstallerに含めるには
* popplerをまるごと`~~~.py`と同じフォルダに入れる。
* pythonのスクリプト内で、popplerフォルダへの環境変数を通す。リンク１を参考に。
* `pdf2image.convert_from_path(outputfile , popplerpath)`を指定する。
`
pdf2image.convert_from_path('path/to/pdf',poppler_path=r"path\to\poppler\bin") 
`

* 以下のコードで、一回コンパイルしてみる。
  * `--add data`で、`~~~.py`から見たpopplerフォルダのの相対パスを指定する
  * `--onefile`はうまくいかないらしいので、`--onedir`にする。
```
pyinstaller --onedir --additional-hooks-dir hooks  --add-data "poppler-21.03.0/*;./poppler"   pdf2pptx_win.py 
```
* ですが、まだこれだとうまくいかない！作成された.exeからみて、`poppler/bin`のフォルダがどこにあるかをみる。
* それにあわせて、環境変数のパスと、`pdf2image.convert_from_path(outputfile , popplerpath)`のpoppler pathを修正する。

* 以下のサイトが参考になった。
* [リンク１](https://stackoverflow.com/questions/66303806/how-do-i-include-poppler-to-pyinstaller-generated-exe-when-using-pdf2image)