# 開発・公開メモ

## 想定リポジトリ構成

```text
document-data-extractor/
├─ main.py
├─ OCR.yaml
├─ DocumentDataExtractor.spec
├─ README.md
├─ LICENSE
├─ requirements.txt
├─ pyproject.toml
├─ CHANGELOG.md
├─ .gitignore
├─ docs/
│  ├─ USAGE.md
│  ├─ DATA_QUALITY.md
│  ├─ LIMITATIONS.md
│  ├─ ARCHITECTURE.md
│  ├─ LICENSES.md
│  ├─ SAMPLE_DATA.md
│  └─ images/
└─ samples/
   ├─ business/
   └─ hobby/
```

## 公開前チェック

- [ ] 最新版プログラムを `main.py` として配置
- [ ] `APP_VERSION` と `pyproject.toml` のversionを一致
- [ ] READMEのversion表記を一致
- [ ] `OCR.yaml` を配置
- [ ] `DocumentDataExtractor.spec` を配置
- [ ] requirements.txtで新規環境からインストール確認
- [ ] 個人情報・会社情報・実在資料が残っていないことを確認
- [ ] ローカル絶対パスがコードやログに残っていないことを確認
- [ ] サンプルデータに「サンプルデータ」と明記
- [ ] Python版でOCR環境テストを確認
- [ ] exe版でOCR環境テストを確認
- [ ] OCRモデル初回ダウンロードを確認
- [ ] PDF直接読み取りを確認
- [ ] OCR必須PDFを確認
- [ ] XLSXを確認
- [ ] CSV UTF-8 / CP932を確認
- [ ] Excel / CSV出力を確認
- [ ] Bronze / Silver Candidate表示を確認
- [ ] Windowsの別PCで可能なら動作確認
- [ ] ライセンス条件を再確認

## バージョン更新

アプリ本体:

```python
APP_VERSION = "x.y.z"
```

同時に更新:

- `pyproject.toml`
- `README.md`
- `CHANGELOG.md`

## exeビルド

公開用exeは `DocumentDataExtractor.spec` を使用してビルドします。

```powershell
python -m PyInstaller --noconfirm --clean DocumentDataExtractor.spec
```

配布時は `dist/DocumentDataExtractor/` フォルダ全体をZIP化します。

exe単体では `_internal` 以下の依存ファイルが不足するため、exeだけを配布しないでください。

## Git管理で除外候補

公開前に `.gitignore` を追加してください。

例:

```text
venv/
.venv/
__pycache__/
*.pyc
*.log
dist/
build/
.paddlex/
.DS_Store
Thumbs.db
```

`DocumentDataExtractor.spec` は公開ビルド手順の一部なので、Git管理対象とします。

ユーザー文書、社内資料、実テストデータ、個人情報を含むログ等をリポジトリへ誤追加しないよう注意してください。

## スクリーンショット

公開用スクリーンショットは `docs/images/` に配置することを推奨します。

例:

```text
docs/images/
├─ main-window.png
├─ ocr-result.png
└─ export-example.png
```

READMEからは相対パスで参照します。

```markdown
![Main window](docs/images/main-window.png)
```
