# ヤフオク!在庫表作成アプリ

ヤフオクに出品中の商品の情報を取得して、在庫数と販売開始価格をまとめたExcelファイルを作成します。


## 必要要件

- Python 3.10以上
- Google Chrome / Chromium

## インストール方法
リポジトリをクローンして、依存ライブラリをインストールしてください。

```sh
$ git clone https://github.com/ecoreuse/ya-stock
$ cd ya-stock
$ pip install -r requirements.txt
```

## アプリ使用方法

```sh
$ python app.py
```

```sh
$ python app.py --help
Usage: app.py [OPTIONS]

  ヤフオク!在庫表作成アプリ

Options:
  -u, --username TEXT    ヤフオク!のユーザー名
  -c, --cookiefile PATH  cookies.jsonのパス
  --costrate FLOAT       オークション開始価格に対する仕入れ価格の割合
  --open-xlsx            作成されたExcelファイルを開く
  --help                 Show this message and exit.
```


## ライセンス
MIT License