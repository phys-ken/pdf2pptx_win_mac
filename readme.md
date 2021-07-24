* クロスコンパイルはできないことが判明
* windowsでも環境変数をチェックしたから、macと同じコードでinstallできる。
* ソースを圧縮して配布しよう。
* winについては、winで作成。

```
pyinstaller --onefile --additional-hooks-dir hooks  main.py       
```