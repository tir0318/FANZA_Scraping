# Name

FANZA_Scraping

Fanza上の作品情報をスクレイピングするプログラム

# Features

作品の画像を保存する
作品のURL,タイトル、女優名、作品時間を取得
取得した情報を指定のExcelファイルにまとめる

# Requirement

* Windows 11
* pip 25
* Python 3.12.3
* requests 2.31.0
* beautifulsoup4 4.12.3
* openpyxl 3.1.5

# Installation

pythonのダウンロードリンク<br>
https://www.python.org/downloads/

```bash
python -m pip install --upgrade pip setuptools
pip install beautifulsoup4
pip install openpyxl
```

# Usage

```bash
git clone https://github.com/tir0318/FANZA_Scraping.git
```
1．プログラムを実行
2．「ファイルを選択して実行」ボタンをクリック
3．URLリストファイルを選択
4．作業用Excelファイルを選択
5．確認ダイアログで「はい」をクリック
6．プログレスバーと状態表示で進行状況を確認

このコードは前のバージョンよりも堅牢で、FANZAサイトの変更にも対応しやすくなっています。また、設定の保存機能によって次回の使用時も前回の設定を引き継ぐことができます。

FANZAのウェブサイト構造が大幅に変わった場合は、セレクターの更新が必要になる可能性があります
スクレイピングの頻度や量によっては、サイト側からのアクセス制限がかかる可能性があります
プロキシやIPローテーションなどの高度な機能は実装していないため、大量の処理には向いていません

# Note

バージョンや環境によって動作しない可能性があるため、最新バージョンにしておくようにお願いします。

# Author

* tir0318
* 個人
