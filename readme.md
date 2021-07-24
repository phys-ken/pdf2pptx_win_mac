* クロスコンパイルはできないことが判明
* windowsでも環境変数をチェックしたから、macと同じコードでinstallできる。
* ソースを圧縮して配布しよう。

```
pyinstaller --onefile --additional-hooks-dir hooks  main.py       
```