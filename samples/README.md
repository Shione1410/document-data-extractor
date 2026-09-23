# サンプルデータパック

このフォルダは、文書データ抽出アプリの動作確認用サンプルを
`business` と `hobby` に分けて整理したものです。

すべて架空の「サンプルデータ」であり、実在の企業・製品・ゲームとは無関係です。

## business

- `business_01_pdf_direct.pdf`
  - 文字レイヤーを持つPDF
  - PDF直読みの確認用

- `business_02_xlsx.xlsx`
  - Excelファイル
  - XLSX直接抽出の確認用

- `business_03_csv.csv`
  - CSVファイル
  - CSV直接抽出の確認用

- `business_04_ocr_required.pdf`
  - 画像だけのPDF
  - OCRが必須

- `business_05_ocr_required_difficult.pdf`
  - 画像だけのPDF
  - 傾き・ノイズあり
  - OCRの難しい例

## hobby

- `hobby_01_pdf_direct.pdf`
  - 文字レイヤーを持つ攻略サイト風PDF
  - PDF直読みの確認用

- `hobby_02_xlsx.xlsx`
  - ゲームアイテム紹介風Excel
  - XLSX直接抽出の確認用

- `hobby_03_csv.csv`
  - 攻略Wiki風アイテム一覧CSV
  - CSV直接抽出の確認用

- `hobby_04_ocr_required.pdf`
  - 攻略サイト保存風の画像PDF
  - OCRが必須

- `hobby_05_ocr_required_difficult.pdf`
  - 薄い文字・傾き・ノイズあり
  - OCRの難しい例

## 想定するデモ

アプリの「フォルダを指定」で `business` または `hobby` を選択すると、
PDF直読み・XLSX・CSV・OCR必須PDFをまとめて一括解析できます。
