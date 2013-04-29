# AutoExRunner

## 概要
AutoExRunnerは、ソフトウェアのインテグレーションテスト(総合テスト)を自動化するための、[AutoIt](http://ja.wikipedia.org/wiki/AutoIt)プログラム集です。  
ソフトウェアテストでは、テスト条件を網羅するために、同じ様なテストを繰り返し行う事があります。  
AutoExRunnerは、同じ様なテストを繰り返し行うテストプログラムを作成するのに役立つ、次の３種類のプログラムを提供します。

* 表計算ソフトから１行ずつ値を読み出すプログラム
* アプリケーションを操作するためのプログラム
* ユーティリティプログラム

## 実行環境
AutoExRunnerを利用するには、以下のソフトウェアをインストールして下さい。

1. [AutoIt v3](http://www.autoitscript.com/site/autoit/downloads/)
2. 次のソフトウェアのいずれか。
  - Microsoft Excel
  - [Apache OpenOffice](http://www.openoffice.org/ja/)
  - [Libre Office](http://ja.libreoffice.org/home/)

## 使用方法
基本的な使用方法は、以下の通りです。

1. 表計算ソフトにテスト条件を記入する。
2. テストプログラム(主にコールバック関数)を記述する。

テストプログラムには、次の様な処理を記述する事になります。

#### Microsoft Excelの場合
    #include "runner\AutoExcelRunner.au3"
    
    Local $handle = AutoExcelRunner_Open("[ファイル名]")
    If IsArray($handle) Then
        AutoExcelRunner_Run($handle, "[シート名]", "CallBackFunc")
        AutoExcelRunner_Close($handle)
    EndIf

    Func CallBackFunc(Const $handle)
        Local $value = AutoExcelRunner_GetString($handle, [項目名])

        ～ アプリケーションを操作するテストコードを記述 ～
    EndFunc

#### Apache OpenOffice もしくは Libre Office の場合
    #include "runner\AutoOpenOfficeRunner.au3"
    
    Local $handle = AutoOpenOfficeRunner_Open("[ファイル名]")
    If IsArray($handle) Then
        AutoOpenOfficeRunner_Run($handle, "[シート名]", "CallBackFunc")
        AutoOpenOfficeRunner_Close($handle)
    EndIf

    Func CallBackFunc(Const $handle)
        Local $value = AutoOpenOfficeRunner_GetString($handle, [項目名])

        ～ アプリケーションを操作するテストコードを記述 ～
    EndFunc

## アプリケーション
次のアプリケーションを操作するためのプログラムを提供しています。

* TeraTerm
* Wireshark
* OpenSSL

## サンプルコード
以下、AutoExRunnerを利用したサンプルコードです。

* [CertificateGenerator](https://github.com/310ken1/AutoExRunner/tree/CertificateGenerator)  
[X.509証明書](http://ja.wikipedia.org/wiki/X.509)を生成するプログラム。  
注意: Windows7 64bit版でしか動作確認をしていないため、他の環境だと動作しないかもしれません。

## ライセンス
ライセンスは、[パブリックドメイン](http://ja.wikipedia.org/wiki/%E3%83%91%E3%83%96%E3%83%AA%E3%83%83%E3%82%AF%E3%83%89%E3%83%A1%E3%82%A4%E3%83%B3)です。

このプロダクトは、OpenSSLツールキットを利用するために、OpenSSLプロジェクトによって開発されたソフトウェアを含みます。
(This product includes software developed by the OpenSSL Project for use in the OpenSSL Toolkit)」
