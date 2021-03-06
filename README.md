# Excel Archive (ezip)

## これはなに
日本のIT業界標準である、Excelにあらゆるファイルを埋め込みメールで送信する。といったプロトコルを単純化するためのソフトウェアです。

## 経緯
前述のとおり、日本のIT業界ではいかにあげるプロトコルが標準的に使用されています。

1. Excelを起動
2. 添付したいファイルをExcelに埋め込み、保存
3. Excelファイルをメールに添付、送信
4. オプションでパスワードを別メールで送付

このプロトコルの問題点は、あるファイルを送信するために一度エクセルを起動する必要がある。ということです。

## このソリューションが解決する問題点
このソフトウェアを使用することで、前述のプロトコルであげる1, 2 を省略することが可能になります。

また、このソフトウェアはExcelに依存していないためExcelがインストールされていないPCでも
標準的プロトコルを使用することができる。というメリットがあります。

## 実行環境
このソフトウェアの実行には以下のソフトウェア環境が必要です。

- Windows Vista, 7, 8, 8.1
- .NET Framework 4.5

## 使用方法
このソフトウェアはコマンドラインから使用します。
次に簡単な使用例を示します。

> ezip -o attached-file.xlsx Secret.png

このコマンドラインを実行することで、Secret.png を埋め込んだattached-file.xlsx を生成することができます。
出力結果のExcelファイルは、Office 2007以降で開くことができます。

また、埋め込むファイルは複数個指定することができ、Windowsで扱えるすべてのファイルに対応しています。

## 制約条件
現行バージョンでは、暗号化Excelファイルを作成することができません。

## その他
author: @tmyt