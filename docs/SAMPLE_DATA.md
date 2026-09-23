# サンプルデータ

公開用サンプルは、実際の利用場面をイメージしやすいように2カテゴリに分けます。

## business

業務利用を想定した架空の資材・仕様データです。

推奨構成:

- PDF直接読み取り
- XLSX
- CSV
- OCR必須PDF
- OCR必須PDF（難しい例）

例:

```text
business/
├─ business_01_pdf_direct.pdf
├─ business_02_xlsx.xlsx
├─ business_03_csv.csv
├─ business_04_ocr_required.pdf
└─ business_05_ocr_required_difficult.pdf
```

## hobby

個人利用・趣味利用を想定した架空ゲームのアイテムデータです。

公式サイト、攻略Wiki、攻略記事などを保存した場合をイメージしたレイアウトを使用します。

```text
hobby/
├─ hobby_01_pdf_direct.pdf
├─ hobby_02_xlsx.xlsx
├─ hobby_03_csv.csv
├─ hobby_04_ocr_required.pdf
└─ hobby_05_ocr_required_difficult.pdf
```

## サンプル作成ルール

- 実在企業名を使用しない
- 実在製品番号を使用しない
- 実在ゲームの文章・画像・固有デザインをコピーしない
- 個人情報を含めない
- 各資料の冒頭に「サンプルデータ」と明記する
- 難しいOCRサンプルは「精度保証用」ではなく制約確認用として扱う

## デモ方法

`business` または `hobby` フォルダを `フォルダを指定` で読み込みます。

これにより、形式が混在していても一括で解析できることを示せます。

### 推奨デモ

1. `business` フォルダを一括追加
2. PDF / XLSX / CSVの直接抽出を確認
3. OCR必須PDFがOCRへ切り替わることを確認
4. 結果を一部修正・分類
5. 必要に応じてキーワード整列
6. Excel / CSVへ出力

スクリーンショット撮影時は、サンプルデータだけを使用し、ローカルユーザー名・社内資料名・個人情報が画面に映り込まないよう注意してください。
