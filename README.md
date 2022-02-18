# python-ImageCompare-ToExcelFile
２つのフォルダを指定して、その中の同名の画像どうしを比較して、その差異の画像とともにExcelファイルに貼り付けて出力する。


## はじめに

このプログラムは、２つのフォルダを指定して、その中の同名の画像どうしを比較して、その差異の画像とともにExcelファイルに貼り付けて出力するプログラムです。
↓こんなExcelファイルを出力します。※サンプル画像は私のスマフォで適当にキャプチャ撮ったものです。他では使わないでくださいね。

![image](https://user-images.githubusercontent.com/64426512/154729605-e5e7006f-bf77-44cf-97b1-fe4bfcbbdbdc.png)

指定したフォルダ内にある、同一名称の画像を左右に表示し、真ん中にその２画像の差異を示す画像を生成して貼り付けています。
差異が無い箇所が、灰色です。


詳しくは、↓Qiitaの記事を参照してください。
［ ※記載中 ］

## 必要なライブラリのインストール

Windowsの場合であれば、**「install_library.bat」** を実行すれば必要なライブラリがインストールされるはずです。
LinuxやMacの場合でしたら、テキストファイルでその中身を見て、実行してみてください。

## プログラムの実行方法

↓のコマンドで動作するはずです。
> python main.py

出力結果は、「save.xlsx」になります。




