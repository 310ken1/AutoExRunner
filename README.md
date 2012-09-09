AutoExRunner
=============

概要
-----
AutoExRunner は、 Excel もしくは OpenOffice に入力されたデータを元に 定型処理を実行するための AutoIt スクリプトです。


実行環境
--------
AutoExRunner を実行するためには、以下のソフトウェアをインストールして下さい。

1. [AutoIt v3](http://www.autoitscript.com/site/autoit/downloads/)
2. 次の何れかのソフトウェアをインストールする。
  - Microsoft Excel
  - [Apache OpenOffice](http://www.openoffice.org/ja/)
  - [Libre Office](http://ja.libreoffice.org/home/)


使用方法
--------
AutoExRunnerTest.au3 を参照して下さい。


ライセンス
---------
ライセンスは、[パブリックドメイン](http://ja.wikipedia.org/wiki/%E3%83%91%E3%83%96%E3%83%AA%E3%83%83%E3%82%AF%E3%83%89%E3%83%A1%E3%82%A4%E3%83%B3)です。

このプロダクトは、OpenSSLツールキットを利用するために、OpenSSLプロジェクトによって開発されたソフトウェアを含みます。
(This product includes software developed by the OpenSSL Project for use in the OpenSSL Toolkit)」


既知の問題
----------
### OpenOfficeCalc.au3
- 稀に "予期しないエラーが発生し、LibreOfficeがクラッシュしました" が発生する。
