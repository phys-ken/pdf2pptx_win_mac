* クロスコンパイルはできないことが判明
* windowsでも環境変数をチェックしたから、macと同じコードでinstallできる。

```
pyinstaller --onefile --additional-hooks-dir hooks  main.py       
```