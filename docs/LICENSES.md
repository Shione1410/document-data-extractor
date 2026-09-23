# 依存ライブラリとライセンス

> この文書は開発上の整理を目的としたもので、法的助言ではありません。公開・再配布・実行ファイル化・商用利用の前に、各プロジェクトの最新ライセンスを必ず確認してください。

## 本プロジェクト

ライセンス:

- GNU Affero General Public License v3.0 or later
- SPDX: `AGPL-3.0-or-later`

## PyMuPDF

本アプリはPDF処理にPyMuPDFを使用します。

PyMuPDF / Artifex製品のオープンソース利用にはGNU AGPLv3系の条件が関係します。AGPL条件に従えない場合は商用ライセンスの検討が必要です。

Official licensing:

https://artifex.com/licensing

## PySide6 / Qt for Python

Qt for Python / PySide6はCommunity EditionではLGPLv3 / GPLv3、または商用ライセンスで提供されています。

Official documentation:

https://doc.qt.io/qtforpython-6/

再配布時にはQt / PySide6のライセンス条件および第三者ライセンス表示を確認してください。

## PaddleOCR

PaddleOCRはApache License 2.0で公開されています。

Official repository:

https://github.com/PaddlePaddle/PaddleOCR

## PaddleX

PaddleOCR 3.xのOCRパイプライン構成で使用します。

Official repository:

https://github.com/PaddlePaddle/PaddleX

## PaddlePaddle

OCR推論エンジンとして使用します。

Official repository:

https://github.com/PaddlePaddle/Paddle

## openpyxl

Excel処理に使用します。

Official project:

https://openpyxl.readthedocs.io/

## python-docx

Word文書処理に使用します。

Official project:

https://python-docx.readthedocs.io/

## PyInstaller

Windows向けexe / onedir配布物の生成に使用します。

Official project:

https://pyinstaller.org/

## 推奨

GitHub公開時には少なくとも以下を行ってください。

1. リポジトリ直下に本プロジェクトの `LICENSE` を置く
2. READMEからライセンスへ案内する
3. 実行ファイルを配布する場合は依存ライブラリのライセンス表示を別途整理する
4. 商用・クローズドソース化を検討する場合はPyMuPDF / Qt等の条件を再確認する
5. 依存ライブラリのバージョン更新時はライセンス条件も再確認する
