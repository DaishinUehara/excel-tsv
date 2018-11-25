# excel-tsv

## 概要

xlsx拡張子のexcelファイルのセル値を読み取り、
タブに区切ったtsvファイルとして出力します。
cell内の以下のコードはそれそれ以下の表のように変換されます。

|cell内の値(変換前)|変換後|
|:---|:---|
|\    |\\\\ |
|改行 |\\n   |
|タブ |\\t |
|タブ |\\t |
|\"  |\\" |

## コマンド

java -jar excel-tsv-1.0-jar-with-dependencies.jar excelフォルダ tsv出力フォルダ

