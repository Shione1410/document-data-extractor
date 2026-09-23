import sys
import csv
import json
import tempfile
import traceback
import warnings
import re
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path
from typing import List, Dict, Any

from PySide6.QtCore import Signal, Qt, QObject, QThread, QTimer
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QFileDialog, QTableWidget, QTableWidgetItem, QLabel, QRadioButton,
    QButtonGroup, QMessageBox, QProgressBar, QHeaderView, QCheckBox,
    QTextEdit, QTabWidget, QSplitter, QComboBox, QLineEdit, QFrame, QGroupBox, QSizePolicy, QScrollArea
)

from openpyxl import load_workbook, Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from docx import Document
import pymupdf as fitz

PADDLE_AVAILABLE = None
PADDLE_OCR_AVAILABLE = None
PADDLE_VERSION = "未読込"
PADDLE_ERROR = ""
PADDLE_OCR_ERROR = ""
PaddleOCR = None


def resource_path(filename):
    if getattr(sys, "frozen", False):
        base_dir = Path(sys._MEIPASS)
    else:
        base_dir = Path(__file__).resolve().parent

    return base_dir / filename


def ensure_ocr_imports(log_callback=None, lifecycle_callback=None):
    """Paddle / PaddleOCR はOCRが必要になった時だけ読み込む。

    起動直後に重いAIライブラリをimportしないことで、GUIの起動を軽くする。
    """
    global PADDLE_AVAILABLE, PADDLE_OCR_AVAILABLE, PADDLE_VERSION
    global PADDLE_ERROR, PADDLE_OCR_ERROR, PaddleOCR

    log = log_callback or (lambda *_: None)
    lifecycle = lifecycle_callback or (lambda *_: None)
    if PADDLE_AVAILABLE is not None and PADDLE_OCR_AVAILABLE is not None:
        if PADDLE_AVAILABLE and PADDLE_OCR_AVAILABLE:
            lifecycle("library", "ready", f"読込済み / Paddle {PADDLE_VERSION}")
        return PADDLE_AVAILABLE and PADDLE_OCR_AVAILABLE

    # Paddle内部のccache警告はOCR利用可否に影響しないので抑制する。
    warnings.filterwarnings(
        "ignore",
        message=r"No ccache found.*",
        category=UserWarning,
    )

    lifecycle("library", "working", "OCRライブラリを読み込み中…")
    log("OCRライブラリを初回読み込み中...")
    try:
        import paddle
        PADDLE_AVAILABLE = True
        PADDLE_VERSION = getattr(paddle, "__version__", "unknown")
    except Exception as e:
        PADDLE_AVAILABLE = False
        PADDLE_ERROR = f"{type(e).__name__}: {e}"
        lifecycle("library", "error", "Paddle読込失敗")
        return False

    try:
        from paddleocr import PaddleOCR as _PaddleOCR
        PaddleOCR = _PaddleOCR
        PADDLE_OCR_AVAILABLE = True
    except Exception as e:
        PADDLE_OCR_AVAILABLE = False
        PADDLE_OCR_ERROR = f"{type(e).__name__}: {e}"
        lifecycle("library", "error", "PaddleOCR読込失敗")
        return False

    log(f"OCRライブラリ読込完了: Paddle {PADDLE_VERSION}")
    lifecycle("library", "ready", f"読込完了 / Paddle {PADDLE_VERSION}")
    return True


APP_VERSION = "1.20.15"
APP_NAME = "文書データ抽出ツール"

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".csv", ".docx", ".pdf"}

OCR_MODES = {
    "fast": {
        "label": "高速",
        "description": "PP-OCRv5 Mobile / 220dpi / 補正OFF",
        "dpi": 220,
        "det_model": "PP-OCRv5_mobile_det",
        "rec_model": "PP-OCRv5_mobile_rec",
        "orientation": False,
        "unwarping": False,
        "textline_orientation": False,
    },
    "standard": {
        "label": "標準（推奨）",
        "description": "PP-OCRv6 Small / 260dpi / 補正OFF",
        "dpi": 260,
        "det_model": "PP-OCRv6_small_det",
        "rec_model": "PP-OCRv6_small_rec",
        "orientation": False,
        "unwarping": False,
        "textline_orientation": False,
    },
    "detail": {
        "label": "詳細解析",
        "description": "PP-OCRv6 Medium / 300dpi / 向き・歪み・文字行補正ON",
        "dpi": 300,
        "det_model": "PP-OCRv6_medium_det",
        "rec_model": "PP-OCRv6_medium_rec",
        "orientation": True,
        "unwarping": True,
        "textline_orientation": True,
    },
}

_OCR_ENGINES = {}
_OCR_ENGINE_ERRORS = {}


def get_ocr_engine(mode="standard", log_callback=None, status_callback=None, lifecycle_callback=None):
    log = log_callback or (lambda *_: None)
    status = status_callback or (lambda *_: None)
    lifecycle = lifecycle_callback or (lambda *_: None)
    config = OCR_MODES[mode]

    if mode in _OCR_ENGINES:
        lifecycle("engine", "ready", f"{config['label']} 起動済み")
        status(f"OCR準備完了: {config['label']} / {config['description']}")
        return _OCR_ENGINES[mode]

    if not ensure_ocr_imports(log, lifecycle):
        return None

    try:
        lifecycle("engine", "working", f"{config['label']} を初期化中…")
        status(f"OCR起動中: {config['label']} / {config['description']}")
        log(f"OCRエンジンを初回初期化中: {config['label']}")
        log(f"  使用モデル: {config['det_model']} + {config['rec_model']}")

        config_path = resource_path("OCR.yaml")
        log(f"  OCR設定: {config_path}")

        engine = PaddleOCR(
            paddlex_config=str(config_path),
            lang="japan",
            text_detection_model_name=config["det_model"],
            text_recognition_model_name=config["rec_model"],
            use_doc_orientation_classify=config["orientation"],
            use_doc_unwarping=config["unwarping"],
            use_textline_orientation=config["textline_orientation"],
        )
        _OCR_ENGINES[mode] = engine
        lifecycle("engine", "ready", f"{config['label']} 準備完了")
        status(f"OCR準備完了: {config['label']} / {config['description']}")
        log(f"OCRエンジン初期化完了: {config['label']}")
        return engine
    except Exception as e:
        err = f"{type(e).__name__}: {e}\n\n{traceback.format_exc()}"
        _OCR_ENGINE_ERRORS[mode] = err
        lifecycle("engine", "error", f"{config['label']} 初期化失敗")
        status(f"OCR起動失敗: {config['label']}")
        return None


class DropTableWidget(QTableWidget):
    files_dropped = Signal(list)

    def __init__(self):
        super().__init__(0, 3)
        self.setHorizontalHeaderLabels(["ファイル名", "種類", "パス"])
        self.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        paths = [Path(url.toLocalFile()) for url in event.mimeData().urls()]
        self.files_dropped.emit(paths)
        event.acceptProposedAction()


class AutoClassifyWorker(QObject):
    progress = Signal(int, int)
    result_ready = Signal(object)
    completed = Signal()
    failed = Signal(str)

    def __init__(self, records, teachers):
        super().__init__()
        self.records = records
        self.teachers = teachers

    @staticmethod
    def _normalize(text):
        return re.sub(r"\s+", "", str(text or "")).casefold()

    @staticmethod
    def _features(text):
        text = str(text or "").strip()
        if not text:
            return {"length": 0.0, "digit": 0.0, "alpha": 0.0, "jp": 0.0, "symbol": 0.0}
        n = max(len(text), 1)
        digits = sum(ch.isdigit() for ch in text)
        alpha = sum(("A" <= ch.upper() <= "Z") for ch in text)
        jp = sum(("ぁ" <= ch <= "ん") or ("ァ" <= ch <= "ヶ") or ("一" <= ch <= "龯") for ch in text)
        symbols = n - sum(ch.isalnum() or (("ぁ" <= ch <= "ん") or ("ァ" <= ch <= "ヶ") or ("一" <= ch <= "龯")) for ch in text)
        return {
            "length": min(n / 20.0, 1.0),
            "digit": digits / n,
            "alpha": alpha / n,
            "jp": jp / n,
            "symbol": max(symbols, 0) / n,
        }

    @staticmethod
    def _feature_similarity(a, b):
        keys = ["length", "digit", "alpha", "jp", "symbol"]
        return max(0.0, 1.0 - sum(abs(a[k] - b[k]) for k in keys) / len(keys))

    def run(self):
        try:
            # 同じ教師文字列は重複比較しない。教師を増やしても無駄な計算を抑える。
            by_role = {}
            exact_roles = {}
            seen = set()
            for role, text in self.teachers:
                norm = self._normalize(text)
                if not norm:
                    continue
                key = (role, norm)
                if key in seen:
                    continue
                seen.add(key)
                sample = {
                    "text": str(text),
                    "norm": norm,
                    "features": self._features(text),
                    "length": len(norm),
                }
                by_role.setdefault(role, []).append(sample)
                exact_roles.setdefault(norm, set()).add(role)

            targets = [
                (i, r) for i, r in enumerate(self.records)
                if r.get("kv_role_source") != "手動"
            ]
            total = max(len(targets), 1)
            updates = {}
            cache = {}

            for pos, (record_index, record) in enumerate(targets, start=1):
                text = str(record.get("text", "")).strip()
                if not text or record.get("manual_blank"):
                    updates[record_index] = ("未判定", "", 0.0)
                else:
                    norm = self._normalize(text)
                    if norm in cache:
                        updates[record_index] = cache[norm]
                    elif norm in exact_roles and len(exact_roles[norm]) == 1:
                        role = next(iter(exact_roles[norm]))
                        result = (role, "自動", 1.0)
                        cache[norm] = result
                        updates[record_index] = result
                    else:
                        f = self._features(text)
                        best_role = "未判定"
                        best_score = 0.0
                        for role, samples in by_role.items():
                            # 全教師に対して高コストな文字列比較はしない。
                            # まず軽い特徴量・文字長で近い教師だけ最大24件に絞る。
                            ranked = []
                            for sample in samples:
                                feature_sim = self._feature_similarity(f, sample["features"])
                                max_len = max(len(norm), sample["length"], 1)
                                length_sim = 1.0 - abs(len(norm) - sample["length"]) / max_len
                                cheap_score = 0.75 * feature_sim + 0.25 * max(0.0, length_sim)
                                ranked.append((cheap_score, sample))
                            ranked.sort(key=lambda x: x[0], reverse=True)

                            role_score = 0.0
                            for _, sample in ranked[:24]:
                                string_sim = SequenceMatcher(None, norm, sample["norm"]).ratio()
                                feature_sim = self._feature_similarity(f, sample["features"])
                                score = 0.65 * string_sim + 0.35 * feature_sim
                                if score > role_score:
                                    role_score = score
                            if role_score > best_score:
                                best_score = role_score
                                best_role = role

                        if best_score >= 0.56:
                            result = (best_role, "自動", round(best_score, 3))
                        else:
                            result = ("未判定", "", round(best_score, 3))
                        cache[norm] = result
                        updates[record_index] = result

                if pos == total or pos % max(1, total // 100) == 0:
                    self.progress.emit(pos, total)

            # 結果を先にメインスレッドへ渡し、その後でワーカー完了を通知する。
            # GUI反映はQThreadが完全停止したあとに開始する。
            self.result_ready.emit(updates)
            self.completed.emit()
        except Exception:
            self.failed.emit(traceback.format_exc())


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.resize(1250, 900)
        self.input_files: List[Path] = []
        self.records: List[Dict[str, Any]] = []
        # 横持ち出力用のキーワードアンカー整列。解析結果そのものは変更しない。
        self.keyword_align_enabled = False

        root = QVBoxLayout(self)
        root.setContentsMargins(18, 16, 18, 16)
        root.setSpacing(10)

        title_label = QLabel(f"{APP_NAME}  v{APP_VERSION}")
        title_label.setObjectName("appTitle")
        root.addWidget(title_label)
        subtitle_label = QLabel(
            "Excel / CSV / Word / PDF / スキャンPDFを解析し、確認・修正後にExcel / CSVへ出力します。"
            "  PDFは上→下・同一行は左→右へ読み順を補正します。"
        )
        subtitle_label.setWordWrap(True)
        subtitle_label.setObjectName("appSubtitle")
        root.addWidget(subtitle_label)

        self.notice_label = QLabel(
            "【初回のみ】選択したOCRモデルをダウンロードするため、1〜3分程度かかる場合があります。"
            "回線速度・PC性能・モードにより前後します。2回目以降は保存済みモデルを利用します。"
        )
        self.notice_label.setWordWrap(True)
        self.notice_label.setStyleSheet(
            "QLabel { background:#fff4cc; border:1px solid #d6a800; padding:8px; font-weight:600; }"
        )
        root.addWidget(self.notice_label)

        self.env_label = QLabel(self.build_environment_status())
        self.env_label.setWordWrap(True)
        root.addWidget(self.env_label)

        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("OCRモード:"))
        self.ocr_mode_group = QButtonGroup(self)
        self.fast_ocr_radio = QRadioButton("高速")
        self.standard_ocr_radio = QRadioButton("標準（推奨）")
        self.detail_ocr_radio = QRadioButton("詳細解析")
        self.standard_ocr_radio.setChecked(True)
        self.ocr_mode_group.addButton(self.fast_ocr_radio)
        self.ocr_mode_group.addButton(self.standard_ocr_radio)
        self.ocr_mode_group.addButton(self.detail_ocr_radio)
        for b in [self.fast_ocr_radio, self.standard_ocr_radio, self.detail_ocr_radio]:
            mode_row.addWidget(b)
        mode_row.addStretch()
        root.addLayout(mode_row)

        self.ocr_status_label = QLabel()
        self.ocr_status_label.setWordWrap(True)
        self.ocr_status_label.setStyleSheet("QLabel { font-weight:600; padding:4px; }")
        root.addWidget(self.ocr_status_label)

        lamp_row = QHBoxLayout()
        self.library_lamp = QLabel("● OCRライブラリ: 未読込")
        self.engine_lamp = QLabel("● OCRエンジン: 未起動")
        self.analysis_lamp = QLabel("● 解析: 未開始")
        self.library_lamp.setStyleSheet("QLabel { color:#777; font-weight:700; padding:4px; }")
        self.engine_lamp.setStyleSheet("QLabel { color:#777; font-weight:700; padding:4px; }")
        self.analysis_lamp.setStyleSheet("QLabel { color:#777; font-weight:700; padding:4px; }")
        lamp_row.addWidget(self.library_lamp)
        lamp_row.addWidget(self.engine_lamp)
        lamp_row.addWidget(self.analysis_lamp)
        lamp_row.addStretch()
        self.toggle_log_button = QPushButton("ログを隠す")
        self.toggle_log_button.setProperty("buttonRole", "secondary")
        self.toggle_log_button.setCheckable(True)
        self.toggle_log_button.setChecked(True)
        self.toggle_log_button.setCursor(Qt.PointingHandCursor)
        self.toggle_log_button.setToolTip("右側のログ / 進捗パネルを表示・非表示にします")
        self.toggle_log_button.setMinimumHeight(30)
        lamp_row.addWidget(self.toggle_log_button)
        root.addLayout(lamp_row)
        self.update_ocr_mode_display()

        file_action_group = QGroupBox("ファイル操作")
        file_action_group.setObjectName("actionGroup")
        button_row = QHBoxLayout(file_action_group)
        button_row.setContentsMargins(12, 10, 12, 10)
        button_row.setSpacing(8)
        self.add_files_button = QPushButton("＋  ファイルを追加")
        self.add_folder_button = QPushButton("＋  フォルダを指定")
        self.clear_button = QPushButton("一覧をクリア")
        self.test_ocr_button = QPushButton("OCR環境テスト")
        self.add_files_button.setProperty("buttonRole", "primary")
        self.add_folder_button.setProperty("buttonRole", "primary")
        self.clear_button.setProperty("buttonRole", "secondary")
        self.test_ocr_button.setProperty("buttonRole", "secondary")
        self.add_files_button.setToolTip("解析対象のExcel / CSV / Word / PDFファイルを追加します")
        self.add_folder_button.setToolTip("フォルダ内の対応ファイルをまとめて追加します")
        self.clear_button.setToolTip("現在のファイル一覧を空にします")
        self.test_ocr_button.setToolTip("PaddleOCRの読み込みとOCRエンジン起動を確認します")
        for b in [self.add_files_button, self.add_folder_button, self.clear_button, self.test_ocr_button]:
            b.setMinimumHeight(36)
            b.setCursor(Qt.PointingHandCursor)
            button_row.addWidget(b)
        button_row.addStretch()
        root.addWidget(file_action_group)

        self.tabs = QTabWidget()

        # ①はページ全体をスクロール可能にする。小さいウィンドウでも
        # テーブルや操作ボタンを潰さず、縦スクロールでアクセスできるようにする。
        input_tab = QWidget()
        input_layout = QVBoxLayout(input_tab)
        input_layout.setContentsMargins(10, 10, 10, 10)
        input_layout.setSpacing(8)
        self.table = DropTableWidget()
        self.table.setMinimumHeight(300)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        input_layout.addWidget(self.table, 1)

        option_row = QHBoxLayout()
        self.force_ocr_check = QCheckBox("PDFを強制OCRする")
        self.force_ocr_check.setToolTip("通常PDFでも画像化して、選択中のOCRモードで解析します。手書き追記がある場合に使用します。")
        option_row.addWidget(self.force_ocr_check)
        option_row.addStretch()
        self.run_button = QPushButton("解析開始")
        self.run_button.setProperty("buttonRole", "primary")
        self.run_button.setMinimumHeight(36)
        self.run_button.setCursor(Qt.PointingHandCursor)
        option_row.addWidget(self.run_button)
        input_layout.addLayout(option_row)

        # ①はファイル一覧テーブル自身にスクロール機能があるため、
        # ページ全体をQScrollAreaで包むと「スクロールの中にスクロール」になり操作感が悪い。
        # 外側スクロールは使わず、一覧テーブルのスクロールだけに統一する。
        input_tab.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tabs.addTab(input_tab, "① ファイル選択・解析")

        result_tab = QWidget()
        result_layout = QVBoxLayout(result_tab)
        result_layout.setContentsMargins(10, 10, 10, 10)
        result_layout.setSpacing(8)
        # ②も各セクションを押し潰さない。必要ならページ自体をスクロールする。
        result_tab.setMinimumWidth(760)
        result_control_row = QHBoxLayout()
        self.review_only_check = QCheckBox("要確認のみ表示")
        self.review_summary_label = QLabel("解析結果: 0件 / 要確認: 0件")
        self.review_summary_label.setStyleSheet("QLabel { font-weight:700; }")
        result_control_row.addWidget(self.review_only_check)
        result_control_row.addWidget(self.review_summary_label)
        self.detail_view_button = QPushButton("詳細情報を表示")
        self.detail_view_button.setCheckable(True)
        self.detail_view_button.setToolTip(
            "ページ/シート、抽出種別、OCR元文字列、判定、判定元などの内部情報を表示します"
        )
        result_control_row.addWidget(self.detail_view_button)
        result_control_row.addStretch()
        self.move_up_button = QPushButton("↑ 上へ")
        self.move_down_button = QPushButton("↓ 下へ")
        self.insert_blank_button = QPushButton("空欄を挿入")
        self.delete_row_button = QPushButton("行を除外")
        for b in [self.move_up_button, self.move_down_button, self.insert_blank_button, self.delete_row_button]:
            result_control_row.addWidget(b)
        result_layout.addLayout(result_control_row)
        help_label = QLabel(
            "通常表示は「ファイル・修正後文字列・OCR信頼度・確認」のみです。"
            " 内部情報は「詳細情報を表示」で確認できます。"
            " 半自動整列では、同じファイル内で行を上下移動し、必要な位置に空欄を挿入してから"
            "横持ち出力すると、その表示順が列順になります。"
        )
        help_label.setWordWrap(True)
        help_label.setStyleSheet("QLabel { color:#555; padding:2px 0 6px 0; }")
        result_layout.addWidget(help_label)

        classify_row = QHBoxLayout()
        self.auto_classify_button = QPushButton("手動判定を教師に残りを自動判定")
        self.reset_auto_classify_button = QPushButton("自動判定だけ未判定に戻す")
        self.classify_summary_label = QLabel("教師: 0件 / 自動判定: 0件")
        self.classify_summary_label.setStyleSheet("QLabel { font-weight:700; }")
        classify_row.addWidget(self.auto_classify_button)
        classify_row.addWidget(self.reset_auto_classify_button)
        classify_row.addWidget(self.classify_summary_label)
        classify_row.addStretch()
        result_layout.addLayout(classify_row)
        classify_help = QLabel(
            "判定プルダウン: 未判定 / Key / Value / 文字列 / 記号 / 不要。手動で変更した行は教師になります。"
            " 教師を増やして再度ボタンを押すと、自動判定を何度でも更新できます。"
        )
        classify_help.setWordWrap(True)
        classify_help.setStyleSheet("QLabel { color:#555; padding:0 0 6px 0; }")
        result_layout.addWidget(classify_help)

        self.result_table = QTableWidget(0, 9)
        self.result_table.setHorizontalHeaderLabels([
            "ファイル", "ページ/シート", "抽出種別", "OCR元文字列", "修正後文字列",
            "OCR信頼度", "確認", "判定", "判定元"
        ])
        # 列幅・行高はユーザーが境界をドラッグして自由に変更できる。
        h_header = self.result_table.horizontalHeader()
        h_header.setSectionResizeMode(QHeaderView.Interactive)
        h_header.setMinimumSectionSize(55)
        h_header.setStretchLastSection(False)

        v_header = self.result_table.verticalHeader()
        v_header.setSectionResizeMode(QHeaderView.Interactive)
        v_header.setMinimumSectionSize(20)
        v_header.setDefaultSectionSize(28)

        # 初期表示幅。OCR文字列2列は以前より少しコンパクトにする。
        initial_widths = [150, 110, 95, 175, 175, 80, 85, 105, 90]
        for col, width in enumerate(initial_widths):
            self.result_table.setColumnWidth(col, width)

        # 通常表示は利用者が日常的に必要とする情報だけに絞る。
        # 詳細情報はボタン操作で展開する。
        self.result_detail_columns = [1, 2, 3, 7, 8]
        for col in self.result_detail_columns:
            self.result_table.setColumnHidden(col, True)

        self.result_table.setAlternatingRowColors(True)

        # スクロールバーはアプリ共通スタイルで統一する。
        self.result_table.verticalScrollBar().setToolTip("つまみをドラッグして上下に移動できます")
        self.result_table.horizontalScrollBar().setToolTip("つまみをドラッグして左右に移動できます")
        self.result_table.setMinimumHeight(300)
        result_layout.addWidget(self.result_table, 1)

        # 出力する列を解析結果の表示項目から選択できる。
        column_select_row = QHBoxLayout()
        column_select_row.addWidget(QLabel("出力列:"))
        self.export_column_checks = {}
        export_column_defs = [
            ("source_file", "ファイル", True),
            ("sheet_or_page", "ページ/シート", False),
            ("element_type", "抽出種別", False),
            ("original_text", "OCR元文字列", False),
            ("text", "修正後文字列", True),
            ("confidence", "OCR信頼度", False),
            ("review_status", "確認", False),
            ("kv_role", "判定", False),
            ("kv_role_source", "判定元", False),
        ]
        for key, label, checked in export_column_defs:
            check = QCheckBox(label)
            check.setChecked(checked)
            self.export_column_checks[key] = check
            column_select_row.addWidget(check)
        column_select_row.addStretch()
        result_layout.addLayout(column_select_row)

        layout_row = QHBoxLayout()
        layout_row.addWidget(QLabel("出力レイアウト:"))
        self.export_layout_group = QButtonGroup(self)
        self.wide_export_radio = QRadioButton("ファイル単位・横持ち（推奨）")
        self.long_export_radio = QRadioButton("レコード単位・縦持ち")
        self.wide_export_radio.setChecked(True)
        self.export_layout_group.addButton(self.wide_export_radio)
        self.export_layout_group.addButton(self.long_export_radio)
        layout_row.addWidget(self.wide_export_radio)
        layout_row.addWidget(self.long_export_radio)
        layout_row.addStretch()
        result_layout.addLayout(layout_row)

        # データ品質レベルを明示する。公開時に「どこまで構造化されたデータか」を誤解させない。
        quality_frame = QFrame()
        quality_frame.setObjectName("qualityPanel")
        quality_layout = QHBoxLayout(quality_frame)
        quality_layout.setContentsMargins(12, 9, 12, 9)
        quality_layout.setSpacing(10)
        quality_title = QLabel("データ品質レベル")
        quality_title.setObjectName("qualityTitle")
        self.quality_badge = QLabel("Bronze")
        self.quality_badge.setObjectName("qualityBadgeBronze")
        self.quality_description = QLabel("未検証の抽出データ。OCR結果・読み順補正・最小限の正規化を含みます。")
        self.quality_description.setWordWrap(True)
        quality_layout.addWidget(quality_title)
        quality_layout.addWidget(self.quality_badge)
        quality_layout.addWidget(self.quality_description, 1)
        result_layout.addWidget(quality_frame)

        quality_help = QLabel(
            "Bronze = 未検証の抽出結果 / Silver Candidate = 構造化・整列を試みた結果（人手確認推奨） / "
            "Silver Verified = 人手確認済み。Silver Verifiedは本アプリが自動認定しません。"
        )
        quality_help.setWordWrap(True)
        quality_help.setObjectName("hintText")
        result_layout.addWidget(quality_help)

        # キーワードアンカー整列。横持ち出力時だけ、指定文字列の位置を各ファイルで揃える。
        keyword_group = QGroupBox("横持ち整列（Silver Candidate）")
        keyword_group.setObjectName("keywordGroup")
        keyword_row = QHBoxLayout(keyword_group)
        keyword_row.setContentsMargins(12, 10, 12, 10)
        keyword_row.setSpacing(8)
        keyword_row.addWidget(QLabel("基準キーワード:"))
        self.keyword_align_edit = QLineEdit()
        self.keyword_align_edit.setPlaceholderText("例: 品番")
        self.keyword_align_edit.setMaximumWidth(220)
        self.keyword_partial_check = QCheckBox("部分一致")
        self.keyword_nfkc_check = QCheckBox("全角/半角を正規化")
        self.keyword_nfkc_check.setChecked(True)
        self.keyword_align_button = QPushButton("キーワード整列を適用")
        self.keyword_align_clear_button = QPushButton("整列を解除")
        self.keyword_align_state_label = QLabel("整列: OFF")
        self.keyword_align_state_label.setObjectName("alignmentState")
        self.keyword_align_button.setProperty("buttonRole", "accent")
        self.keyword_align_clear_button.setProperty("buttonRole", "secondary")
        self.keyword_align_button.setMinimumHeight(34)
        self.keyword_align_clear_button.setMinimumHeight(34)
        self.keyword_align_button.setCursor(Qt.PointingHandCursor)
        self.keyword_align_clear_button.setCursor(Qt.PointingHandCursor)
        self.keyword_align_button.setToolTip("入力したキーワードを基準に、横持ち出力の位置をファイル間で揃えます")
        self.keyword_align_clear_button.setToolTip("キーワード整列を解除し、Bronze相当の並びへ戻します")
        keyword_row.addWidget(self.keyword_align_edit)
        keyword_row.addWidget(self.keyword_partial_check)
        keyword_row.addWidget(self.keyword_nfkc_check)
        keyword_row.addWidget(self.keyword_align_button)
        keyword_row.addWidget(self.keyword_align_clear_button)
        keyword_row.addWidget(self.keyword_align_state_label)
        keyword_row.addStretch()
        result_layout.addWidget(keyword_group)

        keyword_help = QLabel(
            "横持ち出力時のみ使用します。各ファイル内で最初に見つかったキーワード位置を基準に、"
            "前にあるデータを捨てず、必要な空欄を挿入して右方向へそろえます。"
            " キーワードが見つからないファイルは元の並びを保持します。"
        )
        keyword_help.setWordWrap(True)
        keyword_help.setStyleSheet("QLabel { color:#555; padding:0 0 6px 0; }")
        result_layout.addWidget(keyword_help)

        export_row = QHBoxLayout()
        export_row.addWidget(QLabel("出力形式:"))
        self.output_group = QButtonGroup(self)
        self.excel_radio = QRadioButton("Excel (.xlsx)")
        self.csv_radio = QRadioButton("CSV (.csv)")
        self.excel_radio.setChecked(True)
        self.output_group.addButton(self.excel_radio)
        self.output_group.addButton(self.csv_radio)
        export_row.addWidget(self.excel_radio)
        export_row.addWidget(self.csv_radio)
        export_row.addStretch()
        self.export_button = QPushButton("修正結果を出力")
        self.export_button.setProperty("buttonRole", "primary")
        self.export_button.setMinimumHeight(36)
        self.export_button.setCursor(Qt.PointingHandCursor)
        self.export_button.setEnabled(False)
        export_row.addWidget(self.export_button)
        result_layout.addLayout(export_row)

        self.result_scroll = QScrollArea()
        self.result_scroll.setObjectName("tabScrollArea")
        self.result_scroll.setWidgetResizable(True)
        self.result_scroll.setFrameShape(QFrame.NoFrame)
        self.result_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.result_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.result_scroll.setWidget(result_tab)
        self.result_scroll.verticalScrollBar().setToolTip("②ページを上下にスクロールします")
        self.tabs.addTab(self.result_scroll, "② 解析結果・修正・出力")

        # メイン操作領域とログ / 進捗領域を左右に分離する。
        # ログを縦方向に置くと①/②タブの高さを奪うため、公開版UIでは右側ペインとする。
        self.main_splitter = QSplitter(Qt.Horizontal)
        self.main_splitter.setObjectName("mainHorizontalSplitter")
        self.main_splitter.setChildrenCollapsible(False)
        self.main_splitter.setHandleWidth(10)
        self.main_splitter.setOpaqueResize(True)
        self.main_splitter.setToolTip("境界を左右にドラッグすると、作業領域とログ領域の幅を変更できます")

        self.tabs.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tabs.setMinimumWidth(520)
        self.main_splitter.addWidget(self.tabs)

        self.diagnostics_widget = QFrame()
        self.diagnostics_widget.setObjectName("diagnosticsPane")
        self.diagnostics_widget.setFrameShape(QFrame.NoFrame)
        self.diagnostics_widget.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        self.diagnostics_widget.setMinimumWidth(300)
        self.diagnostics_widget.setMaximumWidth(520)
        diagnostics_layout = QVBoxLayout(self.diagnostics_widget)
        diagnostics_layout.setContentsMargins(12, 10, 8, 8)
        diagnostics_layout.setSpacing(8)

        log_header_row = QHBoxLayout()
        log_title = QLabel("ログ / 進捗")
        log_title.setObjectName("logPanelTitle")
        log_header_row.addWidget(log_title)
        log_header_row.addStretch()
        close_log_button = QPushButton("閉じる")
        close_log_button.setProperty("buttonRole", "secondary")
        close_log_button.setCursor(Qt.PointingHandCursor)
        close_log_button.setToolTip("ログパネルを非表示にします。上部の『ログを表示』で再表示できます")
        close_log_button.clicked.connect(lambda: self.toggle_log_button.setChecked(False))
        log_header_row.addWidget(close_log_button)
        diagnostics_layout.addLayout(log_header_row)

        self.progress_label = QLabel("進捗: 待機中")
        self.progress_label.setStyleSheet("QLabel { font-weight:600; }")
        diagnostics_layout.addWidget(self.progress_label)
        self.progress = QProgressBar()
        self.progress.setMinimum(0)
        self.progress.setMaximum(100)
        self.progress.setValue(0)
        self.progress.setFormat("%p%")
        self.progress.setTextVisible(True)
        self.progress.setMinimumHeight(24)
        diagnostics_layout.addWidget(self.progress)

        self.status_label = QLabel("待機中")
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet("QLabel { font-size:11pt; font-weight:700; padding:5px 2px; }")
        diagnostics_layout.addWidget(self.status_label)

        resize_hint = QLabel("↔  境界を左右にドラッグしてログ幅を変更できます")
        resize_hint.setObjectName("resizeHint")
        resize_hint.setWordWrap(True)
        diagnostics_layout.addWidget(resize_hint)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMinimumWidth(260)
        self.log_box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.log_box.setToolTip("OCR処理・エラー・内部処理のログを表示します")
        diagnostics_layout.addWidget(self.log_box, 1)

        self.main_splitter.addWidget(self.diagnostics_widget)
        self.main_splitter.setStretchFactor(0, 5)
        self.main_splitter.setStretchFactor(1, 2)
        self.main_splitter.setCollapsible(0, False)
        self.main_splitter.setCollapsible(1, True)
        self.main_splitter.setSizes([920, 330])
        root.addWidget(self.main_splitter, 1)

        self.setStyleSheet("""
            QWidget {
                font-family: "Segoe UI", "Yu Gothic UI", "Meiryo";
                font-size: 10pt;
                color: #1f2937;
                background: #f7f8fa;
            }
            QLabel#appTitle {
                font-size: 20pt;
                font-weight: 700;
                color: #111827;
                padding: 0 0 2px 0;
            }
            QLabel#appSubtitle, QLabel#hintText {
                color: #5f6b7a;
            }
            QGroupBox#actionGroup {
                background: #ffffff;
                border: 1px solid #d9dee7;
                border-radius: 8px;
                margin-top: 8px;
                font-weight: 700;
            }
            QGroupBox#actionGroup::title, QGroupBox#keywordGroup::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #374151;
            }
            QGroupBox#keywordGroup {
                background: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                margin-top: 8px;
                font-weight: 700;
            }
            QPushButton {
                background: #ffffff;
                border: 1px solid #c8d0dc;
                border-radius: 6px;
                padding: 7px 14px;
                font-weight: 600;
            }
            QPushButton:hover { background: #f1f5f9; border-color: #9aa8ba; }
            QPushButton:pressed { background: #e7edf4; }
            QPushButton:disabled { color: #9ca3af; background: #f3f4f6; border-color: #e5e7eb; }
            QPushButton[buttonRole="primary"] {
                background: #2563eb; color: white; border-color: #1d4ed8;
            }
            QPushButton[buttonRole="primary"]:hover { background: #1d4ed8; }
            QPushButton[buttonRole="accent"] {
                background: #0f766e; color: white; border-color: #0f766e;
            }
            QPushButton[buttonRole="accent"]:hover { background: #0b5f59; }
            QPushButton[buttonRole="secondary"] { background: #ffffff; }
            /* 四角い QCheckBox は v1.18 と同様に Qt / OS 標準描画を使用する。
               解析結果の出力列チェックなどは、標準的なチェックボックス表示へ戻す。 */

            /* 丸い QRadioButton は、選択状態が消えない一般的なラジオボタン表示にする。 */
            QRadioButton {
                spacing: 6px;
                padding: 2px 4px;
            }
            QRadioButton::indicator {
                width: 15px;
                height: 15px;
                border-radius: 8px;
                border: 1px solid #737f8f;
                background: #ffffff;
            }
            QRadioButton::indicator:hover {
                border-color: #2563eb;
            }
            QRadioButton::indicator:checked {
                border: 1px solid #4b5563;
                background: qradialgradient(
                    cx:0.5, cy:0.5, radius:0.5, fx:0.5, fy:0.5,
                    stop:0 #2563eb, stop:0.38 #2563eb,
                    stop:0.43 #ffffff, stop:1 #ffffff
                );
            }
            QRadioButton::indicator:disabled {
                border-color: #b8c0cc;
                background: #f1f3f5;
            }
            QTabWidget::pane { border: 1px solid #d6dce5; background: #ffffff; border-radius: 6px; }
            QScrollArea#tabScrollArea {
                border: none;
                background: #ffffff;
            }
            QScrollArea#tabScrollArea > QWidget > QWidget {
                background: #ffffff;
            }
            QTabBar::tab {
                background: #e9edf3; border: 1px solid #d6dce5; padding: 8px 16px; margin-right: 2px;
            }
            QTabBar::tab:selected { background: #ffffff; font-weight: 700; border-bottom-color: #ffffff; }
            QTableWidget {
                background: #ffffff; gridline-color: #e5e7eb; border: 1px solid #d6dce5; selection-background-color: #dbeafe;
            }
            QHeaderView::section {
                background: #eef2f7; border: none; border-right: 1px solid #d6dce5; border-bottom: 1px solid #d6dce5;
                padding: 6px; font-weight: 700;
            }
            QLineEdit, QComboBox, QTextEdit {
                background: #ffffff; border: 1px solid #cbd5e1; border-radius: 5px; padding: 5px;
            }
            QLineEdit:focus, QComboBox:focus, QTextEdit:focus { border: 1px solid #4f7dd9; }
            QFrame#qualityPanel {
                background: #ffffff; border: 1px solid #d9dee7; border-radius: 7px;
            }
            QLabel#qualityTitle { font-weight: 700; color: #374151; }
            QLabel#qualityBadgeBronze {
                background: #fff2cc; color: #7a4a00; border: 1px solid #d9ad42; border-radius: 10px;
                padding: 4px 10px; font-weight: 700;
            }
            QLabel#qualityBadgeSilver {
                background: #eef2f6; color: #374151; border: 1px solid #aeb8c5; border-radius: 10px;
                padding: 4px 10px; font-weight: 700;
            }
            QLabel#alignmentState { font-weight: 700; color: #667085; padding-left: 4px; }
            QLabel#resizeHint {
                background: #eef4ff; color: #37517e; border: 1px solid #c9d7ef; border-radius: 5px; padding: 5px 8px;
            }
            QFrame#diagnosticsPane {
                background: #f4f6f9;
                border: 1px solid #d7dee8;
                border-radius: 7px;
            }
            QLabel#logPanelTitle {
                font-size: 12pt;
                font-weight: 700;
                color: #26364d;
            }
            QSplitter#mainHorizontalSplitter::handle:horizontal {
                background: #c7d0dc;
                border-left: 1px solid #aeb8c5;
                border-right: 1px solid #aeb8c5;
                margin: 8px 2px;
                border-radius: 4px;
            }
            QSplitter#mainHorizontalSplitter::handle:horizontal:hover {
                background: #7292bf;
                border-left: 1px solid #5f7ea8;
                border-right: 1px solid #5f7ea8;
            }
            QSplitter#mainHorizontalSplitter::handle:horizontal:pressed {
                background: #4f7dd9;
            }
            /* 共通スクロールバー: 端ボタンを消してつまみだけを表示する。 */
            QScrollBar:vertical {
                background: #eef1f5;
                width: 18px;
                margin: 2px 0px 2px 0px;
                border: none;
                border-radius: 8px;
            }
            QScrollBar::handle:vertical {
                background: #9ba8b8;
                min-height: 46px;
                border: none;
                border-radius: 7px;
            }
            QScrollBar::handle:vertical:hover { background: #748399; }
            QScrollBar::handle:vertical:pressed { background: #5f6f85; }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
                background: transparent;
                border: none;
            }
            QScrollBar::add-page:vertical,
            QScrollBar::sub-page:vertical {
                background: transparent;
                border: none;
            }
            QScrollBar:horizontal {
                background: #eef1f5;
                height: 16px;
                margin: 0px 2px 0px 2px;
                border: none;
                border-radius: 7px;
            }
            QScrollBar::handle:horizontal {
                background: #9ba8b8;
                min-width: 46px;
                border: none;
                border-radius: 6px;
            }
            QScrollBar::handle:horizontal:hover { background: #748399; }
            QScrollBar::handle:horizontal:pressed { background: #5f6f85; }
            QScrollBar::add-line:horizontal,
            QScrollBar::sub-line:horizontal {
                width: 0px;
                background: transparent;
                border: none;
            }
            QScrollBar::add-page:horizontal,
            QScrollBar::sub-page:horizontal {
                background: transparent;
                border: none;
            }
        """)
        self.update_quality_level()

        self.add_files_button.clicked.connect(self.add_files)
        self.add_folder_button.clicked.connect(self.add_folder)
        self.clear_button.clicked.connect(self.clear_files)
        self.test_ocr_button.clicked.connect(self.test_ocr_environment)
        self.run_button.clicked.connect(self.run_analysis)
        self.export_button.clicked.connect(self.export_results)
        self.review_only_check.toggled.connect(self.apply_review_filter)
        self.detail_view_button.toggled.connect(self.toggle_result_detail_view)
        self.result_table.cellChanged.connect(self.on_result_cell_changed)
        self.move_up_button.clicked.connect(self.move_selected_row_up)
        self.move_down_button.clicked.connect(self.move_selected_row_down)
        self.insert_blank_button.clicked.connect(self.insert_blank_row)
        self.delete_row_button.clicked.connect(self.delete_selected_row)
        self.auto_classify_button.clicked.connect(self.auto_classify_from_manual_examples)
        self.reset_auto_classify_button.clicked.connect(self.reset_auto_classifications)
        self.keyword_align_button.clicked.connect(self.enable_keyword_alignment)
        self.keyword_align_clear_button.clicked.connect(self.disable_keyword_alignment)
        self.keyword_align_edit.textChanged.connect(self.on_keyword_alignment_setting_changed)
        self.keyword_partial_check.toggled.connect(self.on_keyword_alignment_setting_changed)
        self.keyword_nfkc_check.toggled.connect(self.on_keyword_alignment_setting_changed)
        self.table.files_dropped.connect(self.add_paths)
        self.fast_ocr_radio.toggled.connect(self.update_ocr_mode_display)
        self.standard_ocr_radio.toggled.connect(self.update_ocr_mode_display)
        self.detail_ocr_radio.toggled.connect(self.update_ocr_mode_display)
        self.toggle_log_button.toggled.connect(self.toggle_log_panel)
        self.log_environment_status()



    def toggle_result_detail_view(self, checked):
        """解析結果テーブルの内部メタ情報を必要なときだけ表示する。"""
        for col in getattr(self, "result_detail_columns", []):
            self.result_table.setColumnHidden(col, not checked)

        if checked:
            self.detail_view_button.setText("詳細情報を隠す")
            self.detail_view_button.setToolTip("通常表示に戻します")
        else:
            self.detail_view_button.setText("詳細情報を表示")
            self.detail_view_button.setToolTip(
                "ページ/シート、抽出種別、OCR元文字列、判定、判定元などの内部情報を表示します"
            )


    def update_quality_level(self):
        """現在の出力品質レベルをGUI上で明示する。

        Bronze: 未検証の抽出データ。
        Silver Candidate: キーワード整列など構造化処理を適用した未検証データ。
        Silver Verifiedは人手確認を前提とするため自動認定しない。
        """
        if getattr(self, "keyword_align_enabled", False):
            level = "Silver Candidate"
            desc = "キーワード整列などの構造化処理を適用済み。人手確認を推奨します。"
            badge_name = "qualityBadgeSilver"
        else:
            level = "Bronze"
            desc = "未検証の抽出データ。OCR結果・読み順補正・最小限の正規化を含みます。"
            badge_name = "qualityBadgeBronze"
        if hasattr(self, "quality_badge"):
            self.quality_badge.setText(level)
            self.quality_badge.setObjectName(badge_name)
            self.quality_badge.style().unpolish(self.quality_badge)
            self.quality_badge.style().polish(self.quality_badge)
        if hasattr(self, "quality_description"):
            self.quality_description.setText(desc)
        return level

    def current_quality_level(self):
        return "Silver Candidate" if getattr(self, "keyword_align_enabled", False) else "Bronze"


    def set_lifecycle_state(self, target, state, detail=""):
        colors = {"idle": "#777", "working": "#c48a00", "ready": "#16803a", "error": "#b42318"}
        names = {"library": "OCRライブラリ", "engine": "OCRエンジン", "analysis": "解析"}
        if target == "library":
            label = self.library_lamp
        elif target == "engine":
            label = self.engine_lamp
        else:
            label = self.analysis_lamp
        color = colors.get(state, "#777")
        suffix = f": {detail}" if detail else ""
        label.setText(f"● {names[target]}{suffix}")
        label.setStyleSheet(f"QLabel {{ color:{color}; font-weight:700; padding:4px; }}")
        QApplication.processEvents()

    def update_progress(self, done, total, detail=""):
        total = max(int(total), 1)
        done = min(max(int(done), 0), total)
        percent = int(done * 100 / total)
        self.progress.setRange(0, 100)
        self.progress.setValue(percent)
        self.progress.setFormat(f"{done} / {total}  (%p%)")
        self.progress_label.setText(f"進捗: {detail}" if detail else f"進捗: {done} / {total}")
        QApplication.processEvents()

    def current_ocr_mode(self):
        if self.fast_ocr_radio.isChecked():
            return "fast"
        if self.detail_ocr_radio.isChecked():
            return "detail"
        return "standard"

    def update_ocr_mode_display(self):
        mode = self.current_ocr_mode()
        config = OCR_MODES[mode]
        cached = "起動済み" if mode in _OCR_ENGINES else "未起動"
        self.ocr_status_label.setText(
            f"現在のOCR: {config['label']} | {config['description']} | 状態: {cached}"
        )
        if mode in _OCR_ENGINES:
            self.set_lifecycle_state("engine", "ready", f"{config['label']} 起動済み")
        else:
            self.set_lifecycle_state("engine", "idle", f"{config['label']} 未起動")
        if PADDLE_AVAILABLE is None:
            self.set_lifecycle_state("library", "idle", "未読込")
        elif PADDLE_AVAILABLE and PADDLE_OCR_AVAILABLE:
            self.set_lifecycle_state("library", "ready", f"読込済み / Paddle {PADDLE_VERSION}")
        QApplication.processEvents()

    def set_ocr_status(self, text):
        self.ocr_status_label.setText(text)
        self.status_label.setText(text)
        QApplication.processEvents()

    def toggle_log_panel(self, visible: bool):
        """右側ログ / 進捗パネルを表示・非表示にする。"""
        self.diagnostics_widget.setVisible(visible)
        self.toggle_log_button.setText("ログを隠す" if visible else "ログを表示")
        self.toggle_log_button.setToolTip(
            "右側のログ / 進捗パネルを非表示にします" if visible
            else "右側のログ / 進捗パネルを表示します"
        )
        if visible:
            # 再表示時に極端に細くならないよう、実用的な初期幅へ戻す。
            total = max(self.main_splitter.width(), 1000)
            right = min(360, max(300, total // 4))
            self.main_splitter.setSizes([max(680, total - right), right])

    def build_environment_status(self):
        parts = [f"Python: {sys.version.split()[0]}", f"PyMuPDF: {fitz.VersionBind}"]
        if PADDLE_AVAILABLE is None:
            parts.append("Paddle/OCR: 未読込（必要時に初期化）")
        else:
            parts.append(f"Paddle: OK ({PADDLE_VERSION})" if PADDLE_AVAILABLE else "Paddle: NG")
            parts.append("PaddleOCR: OK / OCR言語: 日本語" if PADDLE_OCR_AVAILABLE else "PaddleOCR: NG")
        return " / ".join(parts)

    def log(self, text):
        self.log_box.append(str(text))
        self.log_box.ensureCursorVisible()
        QApplication.processEvents()

    def log_environment_status(self):
        self.log("=== 起動時環境確認 ===")
        self.log(f"Python: {sys.executable}")
        self.log(f"Python version: {sys.version}")
        self.log(f"PyMuPDF: {fitz.VersionBind}")
        self.log("Paddle / PaddleOCR: 起動時には読み込みません（OCR実行時に初回読込）")
        self.log("OCRモード: 高速 / 標準（推奨） / 詳細解析")
        self.log("初回のOCRモデル取得には1〜3分程度かかる場合があります。")
        self.log("")

    def test_ocr_environment(self):
        mode = self.current_ocr_mode()
        config = OCR_MODES[mode]
        self.log(f"=== OCR環境テスト: {config['label']} ===")
        self.set_ocr_status(f"OCRライブラリ確認中: {config['label']}")
        if not ensure_ocr_imports(self.log, self.set_lifecycle_state):
            if not PADDLE_AVAILABLE:
                self.log(f"失敗: Paddleをimportできません.\n{PADDLE_ERROR}")
                QMessageBox.warning(self, "OCR環境", "Paddleが利用できません。ログを確認してください。")
            else:
                self.log(f"失敗: PaddleOCRをimportできません.\n{PADDLE_OCR_ERROR}")
                QMessageBox.warning(self, "OCR環境", "PaddleOCRが利用できません。ログを確認してください。")
            return
        engine = get_ocr_engine(mode, self.log, self.set_ocr_status, self.set_lifecycle_state)
        if engine is None:
            err = _OCR_ENGINE_ERRORS.get(mode, "原因不明")
            self.log("OCRエンジン初期化失敗:")
            self.log(err)
            QMessageBox.warning(self, "OCR環境", "OCRエンジン初期化に失敗しました。ログを確認してください。")
            return
        self.update_ocr_mode_display()
        self.log(f"OCRエンジン初期化: OK / {config['description']}")
        QMessageBox.information(self, "OCR環境", f"{config['label']} OCR は利用可能です。")

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "ファイルを選択", "", "対応ファイル (*.xlsx *.xlsm *.csv *.docx *.pdf)")
        self.add_paths([Path(p) for p in files])

    def add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "フォルダを選択")
        if folder:
            self.add_paths([Path(folder)])

    def clear_files(self):
        self.input_files.clear()
        self.table.setRowCount(0)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFormat("%p%")
        self.progress_label.setText("進捗: 待機中")
        self.records.clear()
        self.result_table.setRowCount(0)
        self.export_button.setEnabled(False)
        self.review_summary_label.setText("解析結果: 0件 / 要確認: 0件")
        self.set_lifecycle_state("analysis", "idle", "未開始")
        self.status_label.setText("一覧をクリアしました")
        self.log("ファイル一覧と解析結果をクリアしました。")

    def add_paths(self, paths):
        discovered = []
        for path in paths:
            if not path.exists():
                continue
            if path.is_dir():
                discovered += [p for p in path.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS]
            elif path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
                discovered.append(path)
        existing = {p.resolve() for p in self.input_files}
        added = 0
        for file_path in discovered:
            if file_path.resolve() not in existing:
                self.input_files.append(file_path)
                existing.add(file_path.resolve())
                added += 1
        self.refresh_table()
        self.status_label.setText(f"{len(self.input_files)} ファイルを登録")
        self.log(f"{added} ファイル追加。合計 {len(self.input_files)} ファイル。")

    def refresh_table(self):
        self.table.setRowCount(0)
        for file_path in self.input_files:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(file_path.name))
            self.table.setItem(row, 1, QTableWidgetItem(file_path.suffix.lower()))
            self.table.setItem(row, 2, QTableWidgetItem(str(file_path)))

    def run_analysis(self):
        if not self.input_files:
            QMessageBox.warning(self, "確認", "ファイルを追加してください。")
            return

        self.records.clear()
        self.result_table.blockSignals(True)
        self.result_table.setRowCount(0)
        self.result_table.blockSignals(False)
        self.export_button.setEnabled(False)

        total_units = count_work_units(self.input_files)
        completed_units = 0
        self.update_progress(0, total_units, "解析準備中")
        self.set_lifecycle_state("analysis", "working", "解析中")
        errors = []
        self.log("")
        mode = self.current_ocr_mode()
        config = OCR_MODES[mode]
        self.log("=== 解析開始 ===")
        self.log(f"OCRモード: {config['label']} / {config['description']}")
        self.run_button.setEnabled(False)

        for index, file_path in enumerate(self.input_files, start=1):
            self.status_label.setText(f"解析中: {file_path.name}")
            self.log(f"[{index}/{len(self.input_files)}] {file_path.name}")
            QApplication.processEvents()
            try:
                suffix = file_path.suffix.lower()
                if suffix in {".xlsx", ".xlsm"}:
                    result = extract_excel(file_path)
                    completed_units += 1
                    self.update_progress(completed_units, total_units, f"{file_path.name} 完了")
                elif suffix == ".csv":
                    result = extract_csv(file_path)
                    completed_units += 1
                    self.update_progress(completed_units, total_units, f"{file_path.name} 完了")
                elif suffix == ".docx":
                    result = extract_word(file_path)
                    completed_units += 1
                    self.update_progress(completed_units, total_units, f"{file_path.name} 完了")
                elif suffix == ".pdf":
                    base_units = completed_units
                    def page_progress(page_done, page_total, page_detail):
                        self.update_progress(base_units + page_done, total_units, page_detail)
                    result = extract_pdf(
                        file_path,
                        self.force_ocr_check.isChecked(),
                        self.log,
                        mode,
                        self.set_ocr_status,
                        self.set_lifecycle_state,
                        page_progress,
                    )
                    completed_units += get_pdf_page_count(file_path)
                else:
                    result = []

                # PDFだけ、OCRが返した内部順序ではなく「上→下、同じ行は左→右」に安定化する。
                # 複数ファイル間の自動整列は行わず、各ファイル内部の読み順だけを整える。
                if suffix == ".pdf":
                    result = stabilize_pdf_reading_order(result)
                    self.log("  -> 読み順安定化＋近接結合: 行判定 → 行内X順 → 近い断片を1データ化 / ハイフン正規化")

                for record in result:
                    record.setdefault("original_text", record.get("text", ""))
                self.records.extend(result)
                self.log(f"  -> {len(result)} 件取得")
            except Exception as exc:
                errors.append(f"{file_path.name}: {exc}")
                self.log(f"  ERROR: {exc}")
                self.log(traceback.format_exc())

        self.populate_result_table()
        self.update_progress(total_units, total_units, "解析完了・結果を確認してください")
        self.set_lifecycle_state("analysis", "ready", "解析完了")
        self.status_label.setText(f"解析完了: {len(self.records)} 件 / 修正後に出力できます")
        self.log(f"=== 解析完了: {len(self.records)} 件 ===")
        self.run_button.setEnabled(True)
        self.export_button.setEnabled(bool(self.records))
        self.update_ocr_mode_display()
        self.tabs.setCurrentIndex(1)

        message = f"{len(self.records)} 件を解析しました。\n解析結果を確認・修正してから出力してください。"
        if errors:
            message += "\n\n一部エラーがあります。ログを確認してください。"
        QMessageBox.information(self, "解析完了", message)

    def populate_result_table(self):
        self.result_table.blockSignals(True)
        self.result_table.setRowCount(0)
        review_count = 0
        for record_index, record in enumerate(self.records):
            record.setdefault("display_order", record_index)
            record.setdefault("manual_blank", False)
            record.setdefault("excluded", False)
            record.setdefault("kv_role", "未判定")
            record.setdefault("kv_role_source", "")
            row = self.result_table.rowCount()
            self.result_table.insertRow(row)
            original = str(record.get("original_text", record.get("text", "")))
            edited = str(record.get("text", ""))
            confidence = display_ocr_confidence(record.get("confidence"))
            needs_review = False
            try:
                if confidence != "" and float(confidence) < 0.80:
                    needs_review = True
            except Exception:
                pass
            if record.get("element_type", "").startswith("ocr_") and not edited.strip():
                needs_review = True
            if needs_review and not record.get("excluded") and not record.get("manual_blank"):
                review_count += 1

            if record.get("excluded"):
                review_label = "除外"
            elif record.get("manual_blank"):
                review_label = "空欄"
            else:
                review_label = "要確認" if needs_review else "OK"
            values = [
                record.get("source_file", ""), record.get("sheet_or_page", ""),
                display_element_type(record.get("element_type", "")), original, edited, confidence,
                review_label
            ]
            for col, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setData(Qt.UserRole, record_index)
                if col != 4:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.result_table.setItem(row, col, item)

            combo = QComboBox()
            combo.addItems(["未判定", "Key", "Value", "文字列", "記号", "不要"])
            role = str(record.get("kv_role", "未判定"))
            if role not in ["未判定", "Key", "Value", "文字列", "記号", "不要"]:
                role = "未判定"
            combo.setCurrentText(role)
            combo.setProperty("record_index", record_index)
            combo.currentTextChanged.connect(self.on_role_combo_changed)
            self.result_table.setCellWidget(row, 7, combo)

            source = record.get("kv_role_source", "")
            source_item = QTableWidgetItem(str(source))
            source_item.setData(Qt.UserRole, record_index)
            source_item.setFlags(source_item.flags() & ~Qt.ItemIsEditable)
            self.result_table.setItem(row, 8, source_item)

        self.result_table.blockSignals(False)
        self.review_summary_label.setText(f"解析結果: {len(self.records)}件 / 要確認: {review_count}件")
        self.update_classify_summary()
        self.apply_review_filter()

    def on_role_combo_changed(self, role):
        combo = self.sender()
        if combo is None:
            return
        record_index = combo.property("record_index")
        if record_index is None:
            return
        try:
            record = self.records[int(record_index)]
        except Exception:
            return
        record["kv_role"] = role
        # ユーザーがプルダウンを触った時点で、その行は次回判定の教師になる。
        record["kv_role_source"] = "手動" if role != "未判定" else ""
        self.update_classify_summary()
        # 判定元表示だけ即時更新する。
        for row in range(self.result_table.rowCount()):
            item = self.result_table.item(row, 8)
            if item is not None and item.data(Qt.UserRole) == int(record_index):
                item.setText(record["kv_role_source"])
                break

    def update_classify_summary(self):
        manual_count = sum(1 for r in self.records if r.get("kv_role_source") == "手動" and r.get("kv_role") != "未判定")
        auto_count = sum(1 for r in self.records if r.get("kv_role_source") == "自動" and r.get("kv_role") != "未判定")
        self.classify_summary_label.setText(f"教師: {manual_count}件 / 自動判定: {auto_count}件")

    def _classification_features(self, text):
        text = str(text or "").strip()
        if not text:
            return {"length": 0, "digit": 0.0, "alpha": 0.0, "jp": 0.0, "symbol": 0.0}
        n = max(len(text), 1)
        digits = sum(ch.isdigit() for ch in text)
        alpha = sum(("A" <= ch.upper() <= "Z") for ch in text)
        jp = sum(("ぁ" <= ch <= "ん") or ("ァ" <= ch <= "ヶ") or ("一" <= ch <= "龯") for ch in text)
        symbols = n - sum(ch.isalnum() or (("ぁ" <= ch <= "ん") or ("ァ" <= ch <= "ヶ") or ("一" <= ch <= "龯")) for ch in text)
        return {
            "length": min(n / 20.0, 1.0),
            "digit": digits / n,
            "alpha": alpha / n,
            "jp": jp / n,
            "symbol": max(symbols, 0) / n,
        }

    def _feature_similarity(self, a, b):
        keys = ["length", "digit", "alpha", "jp", "symbol"]
        return max(0.0, 1.0 - sum(abs(a[k] - b[k]) for k in keys) / len(keys))

    def auto_classify_from_manual_examples(self):
        teachers = [
            (r.get("kv_role"), str(r.get("text", "")).strip())
            for r in self.records
            if r.get("kv_role_source") == "手動"
            and r.get("kv_role") in {"Key", "Value", "文字列", "記号", "不要"}
            and str(r.get("text", "")).strip()
        ]
        if not teachers:
            QMessageBox.information(self, "自動判定", "まず判定プルダウンで1件以上を Key / Value / 文字列 / 記号 / 不要 に手動設定してください。")
            return

        # 件数が多くてもGUIを固めないよう、判定処理はQThreadへ移す。
        if getattr(self, "auto_classify_thread", None) is not None and self.auto_classify_thread.isRunning():
            return

        self.auto_classify_button.setEnabled(False)
        self.reset_auto_classify_button.setEnabled(False)
        self.status_label.setText("自動判定中…")
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFormat("自動判定 %p%")

        record_snapshot = [dict(r) for r in self.records]
        # v1.15: 判定計算とGUI反映を完全に分離する。
        # 1) workerが結果を返す
        # 2) worker完了 → QThreadを停止
        # 3) QThread.finished 後にだけGUIへの分割反映を開始
        self._auto_classify_result_buffer = None
        self._auto_classify_failed_detail = None
        self.auto_classify_thread = QThread(self)
        self.auto_classify_worker = AutoClassifyWorker(record_snapshot, teachers)
        self.auto_classify_worker.moveToThread(self.auto_classify_thread)

        self.auto_classify_thread.started.connect(self.auto_classify_worker.run)
        self.auto_classify_worker.progress.connect(self.on_auto_classify_progress)
        self.auto_classify_worker.result_ready.connect(self._capture_auto_classify_result)
        self.auto_classify_worker.failed.connect(self._capture_auto_classify_failure)

        # 標準的なQtワーカーパターン。GUI反映はここでは行わない。
        self.auto_classify_worker.completed.connect(self.auto_classify_thread.quit)
        self.auto_classify_worker.failed.connect(self.auto_classify_thread.quit)
        self.auto_classify_worker.completed.connect(self.auto_classify_worker.deleteLater)
        self.auto_classify_worker.failed.connect(self.auto_classify_worker.deleteLater)
        self.auto_classify_thread.finished.connect(self._on_auto_classify_thread_finished)
        self.auto_classify_thread.finished.connect(self.auto_classify_thread.deleteLater)
        self.auto_classify_thread.start()

    def on_auto_classify_progress(self, done, total):
        total = max(int(total), 1)
        percent = int(int(done) * 100 / total)
        self.progress.setValue(percent)
        self.progress.setFormat(f"自動判定 {done} / {total}  (%p%)")
        self.status_label.setText(f"自動判定中… {done} / {total}")

    def _capture_auto_classify_result(self, updates):
        """ワーカー結果は保存だけする。ここではGUIテーブルを触らない。"""
        self._auto_classify_result_buffer = updates
        self.progress.setValue(100)
        self.progress.setFormat("自動判定 計算完了")
        self.status_label.setText("自動判定の計算完了・スレッド終了待ち…")

    def _capture_auto_classify_failure(self, detail):
        """失敗内容も一旦保存し、QThread終了後にGUIへ通知する。"""
        self._auto_classify_failed_detail = detail
        self.status_label.setText("自動判定エラー・スレッド終了待ち…")

    def _on_auto_classify_thread_finished(self):
        """QThreadが完全に停止した後でだけ結果反映を開始する。"""
        # 参照を先に外しておく。次回判定時の isRunning() 判定も安全になる。
        self.auto_classify_worker = None
        self.auto_classify_thread = None

        failed_detail = getattr(self, "_auto_classify_failed_detail", None)
        if failed_detail:
            self._auto_classify_failed_detail = None
            self._auto_classify_result_buffer = None
            self.on_auto_classify_failed(failed_detail)
            return

        updates = getattr(self, "_auto_classify_result_buffer", None)
        self._auto_classify_result_buffer = None
        if updates is None:
            self.on_auto_classify_failed("判定結果を受信できないままQThreadが終了しました。")
            return

        # QThread.finished のシグナル処理を抜けた後のイベントループでGUI反映を開始。
        self.progress.setValue(0)
        self.progress.setFormat("結果反映 %p%")
        self.status_label.setText("自動判定計算完了・画面反映を開始します…")
        QTimer.singleShot(0, lambda u=updates: self._prepare_auto_classify_apply(u))

    def _prepare_auto_classify_apply(self, updates):
        # 判定計算とスレッド終了は完了済み。
        # 既存テーブルを残したまま、結果反映だけを小分けにしてイベントループへ返す。
        row_by_record = {}
        for row in range(self.result_table.rowCount()):
            item = self.result_table.item(row, 0)
            if item is not None:
                idx = item.data(Qt.UserRole)
                if idx is not None:
                    try:
                        row_by_record[int(idx)] = row
                    except Exception:
                        pass

        pending = []
        for index, result in updates.items():
            try:
                index = int(index)
            except Exception:
                continue
            if index < 0 or index >= len(self.records):
                continue
            if self.records[index].get("kv_role_source") == "手動":
                continue
            pending.append((index, result))

        self._classify_apply_pending = pending
        self._classify_apply_row_map = row_by_record
        self._classify_apply_pos = 0
        self._classify_apply_changed = 0
        self._classify_apply_chunk = 50

        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFormat("結果反映 %p%")
        self.status_label.setText(f"自動判定完了・画面へ反映中… 0 / {len(pending)}")
        QTimer.singleShot(1, self._apply_auto_classify_chunk)

    def _apply_auto_classify_chunk(self):
        pending = getattr(self, "_classify_apply_pending", [])
        total = len(pending)
        start = getattr(self, "_classify_apply_pos", 0)
        chunk = max(int(getattr(self, "_classify_apply_chunk", 50)), 1)
        end = min(start + chunk, total)
        row_map = getattr(self, "_classify_apply_row_map", {})

        for pos in range(start, end):
            index, result = pending[pos]
            if index >= len(self.records):
                continue
            record = self.records[index]
            if record.get("kv_role_source") == "手動":
                continue
            role, source, score = result
            record["kv_role"] = role
            record["kv_role_source"] = source
            record["kv_role_score"] = score
            if source == "自動" and role != "未判定":
                self._classify_apply_changed += 1

            row = row_map.get(index)
            if row is None:
                continue

            combo = self.result_table.cellWidget(row, 7)
            if isinstance(combo, QComboBox):
                combo.blockSignals(True)
                combo.setCurrentText(role if role in ["未判定", "Key", "Value", "文字列", "記号", "不要"] else "未判定")
                combo.blockSignals(False)

            source_item = self.result_table.item(row, 8)
            if source_item is not None:
                source_item.setText(str(source))

        self._classify_apply_pos = end
        percent = 100 if total == 0 else int(end * 100 / total)
        self.progress.setValue(percent)
        self.progress.setFormat(f"結果反映 {end} / {max(total, 1)}  (%p%)")
        self.status_label.setText(f"自動判定完了・画面へ反映中… {end} / {total}")

        if end < total:
            QTimer.singleShot(1, self._apply_auto_classify_chunk)
            return

        self._finish_auto_classify_apply()

    def _finish_auto_classify_apply(self):
        changed = int(getattr(self, "_classify_apply_changed", 0))
        self.update_classify_summary()
        self.apply_review_filter()
        self.progress.setValue(100)
        self.progress.setFormat("自動判定 完了")
        self.status_label.setText(f"自動判定完了: {changed}件")
        self.auto_classify_button.setEnabled(True)
        self.reset_auto_classify_button.setEnabled(True)
        teacher_count = sum(1 for r in self.records if r.get("kv_role_source") == "手動" and r.get("kv_role") != "未判定")
        self.log(f"Key/Value/文字列等の自動判定: 教師={teacher_count}件 / 自動判定={changed}件（非同期計算＋分割反映）")

        # 一時データを解放。大量データ時のメモリ保持も避ける。
        self._classify_apply_pending = []
        self._classify_apply_row_map = {}

        # 完了ダイアログも現在の反映処理から切り離す。大量行更新直後に
        # モーダルダイアログを開いてイベントループを詰まらせないためである。
        QTimer.singleShot(50, lambda: QMessageBox.information(
            self, "自動判定完了",
            f"手動教師 {teacher_count} 件を参考に、{changed} 件を自動判定しました。\n"
            "結果は引き続きプルダウンで修正でき、修正した行を教師にして再判定できます。"
        ))

    def on_auto_classify_failed(self, detail):
        self._auto_classify_result_buffer = None
        self._auto_classify_failed_detail = None
        self.auto_classify_button.setEnabled(True)
        self.reset_auto_classify_button.setEnabled(True)
        self.progress.setFormat("自動判定 エラー")
        self.status_label.setText("自動判定でエラーが発生しました")
        self.log("自動判定エラー:")
        self.log(detail)
        QMessageBox.critical(self, "自動判定エラー", "自動判定に失敗しました。ログを確認してください。")

    def reset_auto_classifications(self):
        changed = 0
        for record in self.records:
            if record.get("kv_role_source") == "自動":
                record["kv_role"] = "未判定"
                record["kv_role_source"] = ""
                record.pop("kv_role_score", None)
                changed += 1
        self.populate_result_table()
        self.log(f"自動判定をリセット: {changed}件")

    def on_result_cell_changed(self, row, column):
        if column != 4:
            return
        item = self.result_table.item(row, column)
        if item is None:
            return
        record_index = item.data(Qt.UserRole)
        if record_index is None:
            return
        try:
            self.records[int(record_index)]["text"] = item.text()
        except Exception:
            return

    def _selected_result_row(self):
        rows = sorted({idx.row() for idx in self.result_table.selectedIndexes()})
        return rows[0] if rows else -1

    def _record_index_from_row(self, row):
        if row < 0 or row >= self.result_table.rowCount():
            return None
        item = self.result_table.item(row, 4) or self.result_table.item(row, 0)
        if item is None:
            return None
        value = item.data(Qt.UserRole)
        return int(value) if value is not None else None

    def _same_file_neighbor_row(self, row, direction):
        idx = self._record_index_from_row(row)
        if idx is None:
            return None
        current_file = self.records[idx].get("source_path", "")
        target = row + direction
        if target < 0 or target >= self.result_table.rowCount():
            return None
        tidx = self._record_index_from_row(target)
        if tidx is None or self.records[tidx].get("source_path", "") != current_file:
            return None
        return target

    def _rebuild_records_from_table_order(self):
        ordered_indices = []
        for row in range(self.result_table.rowCount()):
            idx = self._record_index_from_row(row)
            if idx is not None:
                ordered_indices.append(idx)
        if len(ordered_indices) != len(self.records):
            return
        old = list(self.records)
        self.records = [old[i] for i in ordered_indices]
        for i, rec in enumerate(self.records):
            rec["display_order"] = i
        self.populate_result_table()

    def move_selected_row_up(self):
        row = self._selected_result_row()
        target = self._same_file_neighbor_row(row, -1)
        if target is None:
            return
        i = self._record_index_from_row(row)
        j = self._record_index_from_row(target)
        if i is None or j is None:
            return
        self.records[i], self.records[j] = self.records[j], self.records[i]
        self.populate_result_table()
        self.result_table.selectRow(target)

    def move_selected_row_down(self):
        row = self._selected_result_row()
        target = self._same_file_neighbor_row(row, 1)
        if target is None:
            return
        i = self._record_index_from_row(row)
        j = self._record_index_from_row(target)
        if i is None or j is None:
            return
        self.records[i], self.records[j] = self.records[j], self.records[i]
        self.populate_result_table()
        self.result_table.selectRow(target)

    def insert_blank_row(self):
        row = self._selected_result_row()
        if row < 0:
            QMessageBox.information(self, "空欄を挿入", "空欄を挿入したい位置の行を選択してください。")
            return
        idx = self._record_index_from_row(row)
        if idx is None:
            return
        base = self.records[idx]
        blank = base_record(Path(base.get("source_path", base.get("source_file", "blank"))), base.get("source_type", "pdf"))
        blank.update({
            "source_file": base.get("source_file", ""),
            "source_path": base.get("source_path", ""),
            "source_type": base.get("source_type", ""),
            "sheet_or_page": base.get("sheet_or_page", ""),
            "element_type": "manual_blank",
            "original_text": "",
            "text": "",
            "confidence": "",
            "note": "manual blank placeholder",
            "manual_blank": True,
            "excluded": False,
        })
        self.records.insert(idx, blank)
        self.populate_result_table()
        self.result_table.selectRow(row)

    def delete_selected_row(self):
        row = self._selected_result_row()
        if row < 0:
            return
        idx = self._record_index_from_row(row)
        if idx is None:
            return
        self.records[idx]["excluded"] = True
        self.populate_result_table()

    def apply_review_filter(self):
        review_only = self.review_only_check.isChecked()
        for row in range(self.result_table.rowCount()):
            status_item = self.result_table.item(row, 6)
            is_review = status_item is not None and status_item.text() == "要確認"
            self.result_table.setRowHidden(row, review_only and not is_review)

    def normalize_keyword_text(self, value):
        text = str(value or "").strip()
        if self.keyword_nfkc_check.isChecked():
            text = unicodedata.normalize("NFKC", text)
        return text.casefold()

    def keyword_matches_record(self, record, keyword_norm):
        # 修正後文字列を優先し、元OCR文字列も候補にする。
        candidates = [record.get("text", ""), record.get("original_text", "")]
        for value in candidates:
            candidate = self.normalize_keyword_text(value)
            if not candidate:
                continue
            if self.keyword_partial_check.isChecked():
                if keyword_norm in candidate:
                    return True
            elif candidate == keyword_norm:
                return True
        return False

    def enable_keyword_alignment(self):
        keyword = self.keyword_align_edit.text().strip()
        if not keyword:
            QMessageBox.warning(self, "確認", "整列に使うキーワードを入力してください。")
            return
        if not self.records:
            QMessageBox.warning(self, "確認", "先に解析を実行してください。")
            return

        keyword_norm = self.normalize_keyword_text(keyword)
        grouped = {}
        for record in self.records:
            if record.get("excluded"):
                continue
            key = (record.get("source_path", ""), record.get("source_file", ""))
            grouped.setdefault(key, []).append(record)

        found = 0
        positions = []
        for records in grouped.values():
            for idx, record in enumerate(records):
                if self.keyword_matches_record(record, keyword_norm):
                    found += 1
                    positions.append(idx + 1)
                    break

        if not found:
            QMessageBox.warning(self, "キーワード整列", f"『{keyword}』に一致するデータが見つかりませんでした。")
            return

        self.keyword_align_enabled = True
        target = max(positions)
        self.keyword_align_state_label.setText(f"整列: ON / {keyword} / 基準={target}番目 / {found}ファイル")
        self.keyword_align_state_label.setStyleSheet("QLabel { font-weight:700; color:#16803a; }")
        self.update_quality_level()
        self.log(f"キーワード整列ON: keyword={keyword!r} / 対象={found}ファイル / 基準位置={target}")

    def disable_keyword_alignment(self):
        self.keyword_align_enabled = False
        self.keyword_align_state_label.setText("整列: OFF")
        self.keyword_align_state_label.setStyleSheet("QLabel { font-weight:700; color:#666; }")
        self.update_quality_level()
        self.log("キーワード整列OFF")

    def on_keyword_alignment_setting_changed(self, *args):
        # 条件を変えた後に以前の整列条件が残らないよう、安全側で解除する。
        if getattr(self, "keyword_align_enabled", False):
            self.keyword_align_enabled = False
            self.keyword_align_state_label.setText("整列: OFF（条件変更）")
            self.keyword_align_state_label.setStyleSheet("QLabel { font-weight:700; color:#c48a00; }")
            self.update_quality_level()

    def get_selected_export_columns(self):
        return [
            key for key, check in self.export_column_checks.items()
            if check.isChecked()
        ]

    def record_review_status(self, record):
        if record.get("excluded"):
            return "除外"
        if record.get("manual_blank"):
            return "空欄"
        confidence = display_ocr_confidence(record.get("confidence"))
        needs_review = False
        try:
            if confidence != "" and float(confidence) < 0.80:
                needs_review = True
        except Exception:
            pass
        if str(record.get("element_type", "")).startswith("ocr_") and not str(record.get("text", "")).strip():
            needs_review = True
        return "要確認" if needs_review else "OK"

    def build_long_export(self, selected_columns):
        rows = []
        for record in self.records:
            if record.get("excluded"):
                continue
            row = {}
            for key in selected_columns:
                if key == "review_status":
                    row[key] = self.record_review_status(record)
                else:
                    row[key] = record.get(key, "")
            rows.append(row)
        return rows, selected_columns

    def build_wide_export(self, selected_columns):
        # 半自動整列: 解析結果テーブルでユーザーが整えた「ファイル内の表示順」を基本にする。
        # keyword_align_enabled のときだけ、指定キーワードをアンカーとして右方向へ空欄を挿入し、
        # 各ファイルのキーワード位置を共通位置へそろえる。元recordsは変更しない。
        value_columns = [key for key in selected_columns if key != "source_file"]
        grouped = {}
        file_order = []
        for record in self.records:
            file_key = (record.get("source_path", ""), record.get("source_file", ""))
            if file_key not in grouped:
                grouped[file_key] = []
                file_order.append(file_key)
            if record.get("excluded"):
                continue
            grouped[file_key].append(record)

        sequences = {}
        matched_files = 0
        target_anchor = None
        keyword_norm = ""

        if self.keyword_align_enabled:
            keyword_norm = self.normalize_keyword_text(self.keyword_align_edit.text())
            anchor_positions = {}
            for file_key in file_order:
                records = grouped[file_key]
                anchor_idx = None
                for idx, record in enumerate(records):
                    if self.keyword_matches_record(record, keyword_norm):
                        anchor_idx = idx
                        break
                anchor_positions[file_key] = anchor_idx
                if anchor_idx is not None:
                    matched_files += 1

            valid_positions = [p for p in anchor_positions.values() if p is not None]
            if valid_positions:
                # 最も右にあるキーワード位置へ他ファイルを合わせる。
                # こうすれば前方データを捨てず、必要な分だけ先頭へ空欄を追加できる。
                target_anchor = max(valid_positions)

            for file_key in file_order:
                records = list(grouped[file_key])
                anchor_idx = anchor_positions.get(file_key)
                if target_anchor is not None and anchor_idx is not None:
                    pad = max(0, target_anchor - anchor_idx)
                    sequences[file_key] = [None] * pad + records
                else:
                    sequences[file_key] = records
        else:
            for file_key in file_order:
                sequences[file_key] = list(grouped[file_key])

        max_records = max((len(sequences[key]) for key in file_order), default=0)
        headers = ["source_file"]
        for index in range(1, max_records + 1):
            for key in value_columns:
                headers.append(f"{index:03d}_{key}")

        rows = []
        for file_key in file_order:
            records = sequences[file_key]
            row = {"source_file": file_key[1]}
            for index in range(1, max_records + 1):
                record = records[index - 1] if index - 1 < len(records) else None
                for key in value_columns:
                    out_key = f"{index:03d}_{key}"
                    if record is None or record.get("manual_blank"):
                        row[out_key] = ""
                    elif key == "review_status":
                        row[out_key] = self.record_review_status(record)
                    else:
                        row[out_key] = record.get(key, "")
            rows.append(row)

        if self.keyword_align_enabled and target_anchor is not None:
            self.log(
                f"横持ち出力: キーワード整列適用 / keyword={self.keyword_align_edit.text()!r} "
                f"/ 一致={matched_files}/{len(file_order)}ファイル / 基準列={target_anchor + 1}"
            )
        return rows, headers

    def export_results(self):
        if not self.records:
            QMessageBox.warning(self, "確認", "先に解析を実行してください。")
            return

        selected_columns = self.get_selected_export_columns()
        if not selected_columns:
            QMessageBox.warning(self, "確認", "出力する列を1つ以上選択してください。")
            return

        if self.wide_export_radio.isChecked():
            export_rows, export_fields = self.build_wide_export(selected_columns)
            layout_name = "ファイル単位・横持ち" + (" / キーワード整列" if self.keyword_align_enabled else "")
        else:
            export_rows, export_fields = self.build_long_export(selected_columns)
            layout_name = "レコード単位・縦持ち"

        ext = ".xlsx" if self.excel_radio.isChecked() else ".csv"
        file_filter = "Excel (*.xlsx)" if ext == ".xlsx" else "CSV (*.csv)"

        dialog = QFileDialog(self, "修正結果の保存先を選択")
        dialog.setAcceptMode(QFileDialog.AcceptSave)
        dialog.setFileMode(QFileDialog.AnyFile)
        dialog.setNameFilter(file_filter)
        suffix = "wide" if self.wide_export_radio.isChecked() else "long"
        dialog.selectFile(f"extracted_data_{suffix}{ext}")
        dialog.setOption(QFileDialog.DontUseNativeDialog, True)

        # Qtの非ネイティブ保存ダイアログは環境によって標準ラベルが
        # 英語表示になるため、公開版では主要ラベルを明示的に日本語化する。
        dialog.setLabelText(QFileDialog.LookIn, "保存先:")
        dialog.setLabelText(QFileDialog.FileName, "ファイル名:")
        dialog.setLabelText(QFileDialog.FileType, "ファイルの種類:")
        dialog.setLabelText(QFileDialog.Accept, "保存")
        dialog.setLabelText(QFileDialog.Reject, "キャンセル")
        if not dialog.exec():
            return
        selected_files = dialog.selectedFiles()
        if not selected_files:
            return
        output_path = selected_files[0]
        if not output_path.lower().endswith(ext):
            output_path += ext

        try:
            if self.excel_radio.isChecked():
                metadata = {
                    "quality_level": self.current_quality_level(),
                    "quality_description": self.quality_description.text(),
                    "keyword_alignment": "ON" if self.keyword_align_enabled else "OFF",
                    "keyword": self.keyword_align_edit.text().strip() if self.keyword_align_enabled else "",
                    "note": "Silver Verifiedは人手確認済みを意味するため、本アプリは自動認定しません。",
                }
                export_xlsx_custom(export_rows, export_fields, Path(output_path), metadata=metadata)
            else:
                export_csv_custom(export_rows, export_fields, Path(output_path))
        except Exception as exc:
            self.log("出力エラー:")
            self.log(traceback.format_exc())
            QMessageBox.critical(self, "出力エラー", str(exc))
            return

        self.status_label.setText(f"出力完了: {output_path}")
        quality_level = self.current_quality_level()
        self.log(f"出力完了: {output_path} / {layout_name} / {quality_level} / {len(export_fields)}列")
        QMessageBox.information(
            self, "出力完了",
            f"修正済み解析結果を出力しました。\n\n{layout_name}\n品質レベル: {quality_level}\n"
            f"{len(export_rows)}行 × {len(export_fields)}列\n\n{output_path}"
        )


def _as_float(value):
    try:
        if value == "" or value is None:
            return None
        return float(value)
    except Exception:
        return None


def normalize_ocr_hyphens(text):
    """OCRで揺れやすいハイフン系文字をASCII '-' に統一する。"""
    if text is None:
        return ""
    value = str(text)
    hyphens = {
        "‐": "-", "‑": "-", "‒": "-", "–": "-", "—": "-", "―": "-",
        "−": "-", "﹘": "-", "﹣": "-", "－": "-", "ー": "-",
    }
    return "".join(hyphens.get(ch, ch) for ch in value)


def _probable_missing_hyphen(left, right):
    """英字ブロック＋数字ブロック等で、ハイフン欠落らしいケースだけ保守的に判定する。"""
    left = normalize_ocr_hyphens(left).strip()
    right = normalize_ocr_hyphens(right).strip()
    if not left or not right:
        return False
    if not re.fullmatch(r"[A-Za-z0-9]+", left) or not re.fullmatch(r"[A-Za-z0-9]+", right):
        return False
    left_alpha = left.isalpha()
    left_digit = left.isdigit()
    right_alpha = right.isalpha()
    right_digit = right.isdigit()
    return (left_alpha and right_digit) or (left_digit and right_alpha)


def _text_profile(text):
    """文字列を大まかな文字種へ分類する。結合判断用であり、意味分類ではない。"""
    text = normalize_ocr_hyphens(text).strip()
    if not text:
        return "empty"
    if re.fullmatch(r"[-+×xX*/.=,:：;；()（）\[\]{}<>＜＞△▲▽▼○●□■◇◆※#%％℃°φΦ]+", text):
        return "symbol"
    if re.fullmatch(r"[0-9０-９.,]+", text):
        return "number"
    if re.fullmatch(r"[A-Za-zＡ-Ｚａ-ｚ]+", text):
        return "alpha"
    if re.fullmatch(r"[A-Za-zＡ-Ｚａ-ｚ0-9０-９_-]+", text):
        return "alnum"
    # CJK / kana を1文字でも含むか。
    if re.search(r"[\u3040-\u30ff\u3400-\u9fff]", text):
        # 日本語に英数字が多く混在する場合も日本語中心として扱う。
        return "jp"
    return "other"


def _looks_like_unit(text):
    text = normalize_ocr_hyphens(text).strip().lower()
    units = {
        "mm", "cm", "m", "km", "g", "kg", "mg", "l", "ml", "v", "a", "w", "kw",
        "hz", "mhz", "ghz", "pa", "kpa", "mpa", "n", "kn", "℃", "°c", "%", "％",
        "個", "本", "枚", "台", "箱", "回", "日", "月", "年", "秒", "分", "時間",
    }
    return text in units


def _looks_key_like(text):
    """汎用的な『見出しっぽさ』。特定帳票辞書には依存しない。"""
    text = normalize_ocr_hyphens(text).strip()
    if not text:
        return False
    profile = _text_profile(text)
    if profile != "jp":
        return False
    # 長文は見出しより説明文の可能性が高い。
    if len(text) > 18:
        return False
    # 末尾が説明文的な句読点なら見出し扱いしにくい。
    if text.endswith(("。", "、", ".", "!", "！", "?", "？")):
        return False
    return True


def _merge_score(left, right, gap_ratio):
    """距離＋文字種から、隣接OCR断片を1データにまとめる妥当性を点数化する。

    正の値ほど結合寄り、負の値ほど分離寄り。
    特定業界の語彙辞書には依存せず、文字種と一般的な表記パターンだけを見る。
    """
    left = normalize_ocr_hyphens(left).strip()
    right = normalize_ocr_hyphens(right).strip()
    lp = _text_profile(left)
    rp = _text_profile(right)
    score = 0.0

    # まず距離。非常に近ければ強く結合寄り。
    if gap_ratio <= 0.22:
        score += 3.0
    elif gap_ratio <= 0.45:
        score += 2.0
    elif gap_ratio <= 0.75:
        score += 1.0
    elif gap_ratio > 1.05:
        score -= 3.0

    # 同種文字は分割OCRの可能性が高い。
    if lp == rp and lp in {"jp", "alpha", "alnum", "number"}:
        score += 2.0

    # 日本語の短い断片同士は語の途中で切れた可能性がある。
    if lp == "jp" and rp == "jp" and len(left) <= 10 and len(right) <= 10:
        score += 1.5

    # 型番・規格らしい英数字の連続。
    if (lp in {"alpha", "alnum"} and rp in {"number", "alnum"}) or \
       (lp == "number" and rp in {"alpha", "alnum"}):
        score += 1.5

    # ハイフン・寸法記号等は左右をつなぐ記号になりやすい。
    if right in {"-", "×", "x", "X", "/", "."} or left.endswith(("-", "×", "x", "X", "/", ".")):
        score += 3.0
    if left in {"-", "×", "x", "X", "/", "."}:
        score += 3.0

    # 数字＋単位は1データとして扱いやすい。
    if lp == "number" and _looks_like_unit(right):
        score += 3.0

    # 汎用帳票で頻出する『短い日本語見出し + 純数字』は分離を優先。
    if _looks_key_like(left) and rp == "number" and not _looks_like_unit(right):
        score -= 4.0

    # 長めの説明文 + 数字も分ける。
    if lp == "jp" and len(left) >= 10 and rp == "number":
        score -= 3.0

    # 記号だけの断片は、接続記号以外は独立要素の可能性がある。
    if rp == "symbol" and right not in {"-", "×", "x", "X", "/", "."}:
        score -= 1.5

    return score


def _merge_separator(left, right, gap_ratio):
    """結合する場合の間に入れる文字を決める。"""
    left_n = normalize_ocr_hyphens(left)
    right_n = normalize_ocr_hyphens(right)
    if left_n.endswith("-") or right_n.startswith("-"):
        return ""
    if right_n in {"-", "×", "x", "X", "/", "."} or left_n.endswith(("×", "x", "X", "/", ".")):
        return ""
    if _probable_missing_hyphen(left_n, right_n) and gap_ratio <= 0.78:
        return "-"
    # 日本語/英字がほぼ接触している場合は空白なし、それ以外は見た目の空白を残す。
    if gap_ratio <= 0.30:
        return ""
    return " "


def _merge_ocr_line_items(items, join_gap_factor=1.15, merge_threshold=2.5):
    """同じ行のOCR断片を、距離＋文字種スコアで結合する。

    v1.19の『距離だけで近ければ結合』を改め、
    日本語/英数字/数値/記号の組み合わせも見る。
    元OCR文字列は original_text に保持する。
    """
    if not items:
        return []

    merged = []
    current = None

    def start_item(item):
        rec = dict(item["record"])
        raw = str(rec.get("original_text", rec.get("text", "")))
        normalized = normalize_ocr_hyphens(rec.get("text", ""))
        rec["original_text"] = raw
        rec["text"] = normalized
        rec["x0"] = round(item["x0"], 2)
        rec["y0"] = round(item["y0"], 2)
        rec["x1"] = round(item["x1"], 2)
        rec["y1"] = round(item["y1"], 2)
        return {
            "record": rec,
            "x0": item["x0"], "x1": item["x1"],
            "y0": item["y0"], "y1": item["y1"],
            "height": item["height"],
        }

    for item in items:
        rec = item["record"]
        if str(rec.get("element_type", "")) != "ocr_text":
            if current is not None:
                merged.append(current["record"])
                current = None
            plain = dict(rec)
            plain.setdefault("original_text", plain.get("text", ""))
            plain["text"] = normalize_ocr_hyphens(plain.get("text", ""))
            merged.append(plain)
            continue

        if current is None:
            current = start_item(item)
            continue

        cur = current["record"]
        gap = item["x0"] - current["x1"]
        reference_height = max(0.1, (current["height"] + item["height"]) / 2.0)
        gap_ratio = max(0.0, gap) / reference_height

        # 明らかに遠ければスコア計算前に分離。
        if gap > max(2.0, reference_height * join_gap_factor):
            merged.append(cur)
            current = start_item(item)
            continue

        left = str(cur.get("text", ""))
        right_raw = str(rec.get("original_text", rec.get("text", "")))
        right = normalize_ocr_hyphens(rec.get("text", ""))
        score = _merge_score(left, right, gap_ratio)

        if score < merge_threshold:
            merged.append(cur)
            current = start_item(item)
            continue

        sep = _merge_separator(left, right, gap_ratio)
        cur["text"] = normalize_ocr_hyphens(left + sep + right)
        cur["original_text"] = (str(cur.get("original_text", "")).rstrip() + " " + right_raw.lstrip()).strip()

        current["x1"] = max(current["x1"], item["x1"])
        current["y0"] = min(current["y0"], item["y0"])
        current["y1"] = max(current["y1"], item["y1"])
        current["height"] = max(current["y1"] - current["y0"], 0.1)
        cur["x0"] = round(current["x0"], 2)
        cur["y0"] = round(current["y0"], 2)
        cur["x1"] = round(current["x1"], 2)
        cur["y1"] = round(current["y1"], 2)

        try:
            scores = [float(v) for v in (cur.get("confidence", ""), rec.get("confidence", "")) if v != ""]
            if scores:
                cur["confidence"] = round(min(scores), 4)
        except Exception:
            pass
        note = str(cur.get("note", ""))
        cur["note"] = (note + f" / smart-merged(score={score:.1f})").strip(" /")

    if current is not None:
        merged.append(current["record"])
    return merged

def stabilize_pdf_reading_order(records, min_overlap_ratio=0.42, center_factor=0.65):
    """PDFレコードを自然な読み順へ並べ、同じ行の近接OCR断片を1データ化する。

    1. 文字ボックスの高さ・縦重なりで同じ行を適応判定
    2. 行を上→下、行内を左→右へ並べる
    3. 同じ行で距離が近いOCR断片だけを結合
    4. ハイフン類を '-' に統一し、保守的な条件で欠落ハイフンを補う

    元OCR文字列は original_text に保持する。
    """
    if not records:
        return records

    indexed = list(enumerate(records))

    def page_key(record):
        value = record.get("sheet_or_page", "")
        try:
            return (0, int(value))
        except Exception:
            return (1, str(value))

    grouped = {}
    page_order = []
    for original_index, record in indexed:
        key = page_key(record)
        if key not in grouped:
            grouped[key] = []
            page_order.append(key)
        grouped[key].append((original_index, record))

    ordered = []

    for key in sorted(page_order):
        page_items = grouped[key]
        positioned = []
        unpositioned = []

        for original_index, record in page_items:
            x0 = _as_float(record.get("x0"))
            y0 = _as_float(record.get("y0"))
            x1 = _as_float(record.get("x1"))
            y1 = _as_float(record.get("y1"))

            if None in (x0, y0, x1, y1) or y1 <= y0:
                unpositioned.append((original_index, record))
                continue

            height = max(y1 - y0, 0.1)
            center_y = (y0 + y1) / 2.0
            positioned.append({
                "original_index": original_index,
                "record": record,
                "x0": x0, "x1": x1,
                "y0": y0, "y1": y1,
                "height": height,
                "center_y": center_y,
            })

        positioned.sort(key=lambda item: (item["center_y"], item["x0"], item["original_index"]))

        lines = []
        for item in positioned:
            best_line = None
            best_score = None
            for line in lines:
                line_y0 = line["y0"]
                line_y1 = line["y1"]
                line_height = max(line_y1 - line_y0, 0.1)
                line_center = (line_y0 + line_y1) / 2.0
                overlap = max(0.0, min(item["y1"], line_y1) - max(item["y0"], line_y0))
                overlap_ratio = overlap / max(min(item["height"], line_height), 0.1)
                center_distance = abs(item["center_y"] - line_center)
                center_limit = max(3.0, center_factor * max(item["height"], line_height))
                if overlap_ratio >= min_overlap_ratio or center_distance <= center_limit:
                    score = (overlap_ratio, -center_distance)
                    if best_score is None or score > best_score:
                        best_score = score
                        best_line = line

            if best_line is None:
                lines.append({"items": [item], "y0": item["y0"], "y1": item["y1"]})
            else:
                best_line["items"].append(item)
                ys0 = sorted(x["y0"] for x in best_line["items"])
                ys1 = sorted(x["y1"] for x in best_line["items"])
                mid = len(ys0) // 2
                if len(ys0) % 2:
                    best_line["y0"] = ys0[mid]
                    best_line["y1"] = ys1[mid]
                else:
                    best_line["y0"] = (ys0[mid - 1] + ys0[mid]) / 2.0
                    best_line["y1"] = (ys1[mid - 1] + ys1[mid]) / 2.0

        lines.sort(key=lambda line: ((line["y0"] + line["y1"]) / 2.0, line["y0"]))
        for line in lines:
            line["items"].sort(key=lambda item: (item["x0"], item["center_y"], item["original_index"]))
            ordered.extend(_merge_ocr_line_items(line["items"]))

        unpositioned.sort(key=lambda item: item[0])
        for _, record in unpositioned:
            rec = dict(record)
            rec.setdefault("original_text", rec.get("text", ""))
            rec["text"] = normalize_ocr_hyphens(rec.get("text", ""))
            ordered.append(rec)

    return ordered

def get_pdf_page_count(file_path):
    try:
        with fitz.open(file_path) as doc:
            return max(len(doc), 1)
    except Exception:
        return 1


def count_work_units(file_paths):
    total = 0
    for file_path in file_paths:
        if file_path.suffix.lower() == ".pdf":
            total += get_pdf_page_count(file_path)
        else:
            total += 1
    return max(total, 1)


def base_record(file_path: Path, source_type: str) -> Dict[str, Any]:
    return {"source_file": file_path.name, "source_path": str(file_path), "source_type": source_type,
            "sheet_or_page": "", "element_type": "", "text": "", "x0": "", "y0": "", "x1": "", "y1": "",
            "row": "", "column": "", "table_id": "", "confidence": "", "ocr_engine": "", "note": ""}


def extract_excel(file_path):
    records = []
    wb = load_workbook(file_path, data_only=False, read_only=False)
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    r = base_record(file_path, "excel")
                    r.update({"sheet_or_page": sheet.title, "element_type": "cell", "text": str(cell.value),
                              "row": cell.row, "column": cell.column, "table_id": f"sheet:{sheet.title}", "note": cell.coordinate})
                    records.append(r)
    return records



def extract_csv(file_path):
    """CSVをExcelと同様にセル単位で抽出する。

    日本語Windows環境でよく使われるUTF-8(BOM有無)とCP932を順に試す。
    空セルは記録せず、行・列番号を保持する。
    """
    records = []
    rows = None
    used_encoding = None
    last_error = None

    for encoding in ("utf-8-sig", "utf-8", "cp932"):
        try:
            with open(file_path, "r", encoding=encoding, newline="") as f:
                rows = list(csv.reader(f))
            used_encoding = encoding
            break
        except UnicodeDecodeError as exc:
            last_error = exc

    if rows is None:
        raise last_error or UnicodeDecodeError("csv", b"", 0, 1, "CSVの文字コードを判定できません")

    for row_index, row_values in enumerate(rows, start=1):
        for col_index, value in enumerate(row_values, start=1):
            text_value = str(value)
            if not text_value.strip():
                continue
            r = base_record(file_path, "csv")
            r.update({
                "sheet_or_page": "csv",
                "element_type": "cell",
                "text": text_value,
                "row": row_index,
                "column": col_index,
                "table_id": "csv:main",
                "note": f"R{row_index}C{col_index} / encoding:{used_encoding}",
            })
            records.append(r)
    return records


def extract_word(file_path):
    records = []
    doc = Document(file_path)
    for i, p in enumerate(doc.paragraphs, start=1):
        text = p.text.strip()
        if text:
            r = base_record(file_path, "word")
            r.update({"sheet_or_page": "document", "element_type": "paragraph", "text": text, "row": i, "note": f"paragraph:{i}"})
            records.append(r)
    for ti, table in enumerate(doc.tables, start=1):
        for ri, row in enumerate(table.rows, start=1):
            for ci, cell in enumerate(row.cells, start=1):
                text = cell.text.strip()
                if text:
                    r = base_record(file_path, "word")
                    r.update({"sheet_or_page": "document", "element_type": "table_cell", "text": text,
                              "row": ri, "column": ci, "table_id": f"table:{ti}"})
                    records.append(r)
    return records


def make_ocr_error_record(file_path, page_index, element_type, note):
    r = base_record(file_path, "pdf")
    r.update({"sheet_or_page": page_index, "element_type": element_type, "ocr_engine": "PaddleOCR", "note": note})
    return r


def extract_pdf(file_path, force_ocr=False, log_callback=None, ocr_mode="standard", status_callback=None, lifecycle_callback=None, progress_callback=None):
    records = []
    doc = fitz.open(file_path)
    log = log_callback or (lambda *_: None)
    total_pages = max(len(doc), 1)
    for page_index, page in enumerate(doc, start=1):
        native_words = page.get_text("words")
        native_text_length = sum(len(str(w[4]).strip()) for w in native_words)
        log(f"  page {page_index}: native words={len(native_words)}, text chars={native_text_length}")
        should_ocr = force_ocr or native_text_length < 20
        if not should_ocr:
            log(f"  page {page_index}: 通常PDFとして解析")
            records.extend(extract_native_pdf_page(file_path, page, page_index))
            if progress_callback:
                progress_callback(page_index, total_pages, f"{file_path.name}: page {page_index}/{total_pages} 完了（通常PDF）")
            continue
        log(f"  page {page_index}: スキャンPDFと判定 → OCRへ")
        if not ensure_ocr_imports(log, lifecycle_callback):
            if not PADDLE_AVAILABLE:
                log(f"  OCR不可: Paddle import失敗: {PADDLE_ERROR}")
                records.append(make_ocr_error_record(file_path, page_index, "ocr_unavailable", f"Paddle import失敗: {PADDLE_ERROR}"))
            else:
                log(f"  OCR不可: PaddleOCR import失敗: {PADDLE_OCR_ERROR}")
                records.append(make_ocr_error_record(file_path, page_index, "ocr_unavailable", f"PaddleOCR import失敗: {PADDLE_OCR_ERROR}"))
            if progress_callback:
                progress_callback(page_index, total_pages, f"{file_path.name}: page {page_index}/{total_pages} OCR利用不可")
            continue
        config = OCR_MODES[ocr_mode]
        engine = get_ocr_engine(ocr_mode, log, status_callback, lifecycle_callback)
        if engine is None:
            err = _OCR_ENGINE_ERRORS.get(ocr_mode, "原因不明")
            log("  OCRエンジン初期化失敗")
            log(err)
            records.append(make_ocr_error_record(file_path, page_index, "ocr_init_error", err))
            if progress_callback:
                progress_callback(page_index, total_pages, f"{file_path.name}: page {page_index}/{total_pages} OCR初期化失敗")
            continue
        try:
            if status_callback:
                status_callback(f"解析中: page {page_index} / {config['label']} / {config['description']}")
            page_records = extract_ocr_pdf_page(
                file_path, page, page_index, engine, log_callback, ocr_mode
            )
            if not page_records:
                log(f"  page {page_index}: OCRは実行されたが文字0件")
                records.append(make_ocr_error_record(file_path, page_index, "ocr_no_text", "OCR実行済みだが認識文字0件"))
            else:
                records.extend(page_records)
        except Exception as e:
            log(f"  OCR処理例外: {e}")
            log(traceback.format_exc())
            records.append(make_ocr_error_record(file_path, page_index, "ocr_runtime_error", f"{type(e).__name__}: {e}"))
        if progress_callback:
            progress_callback(page_index, total_pages, f"{file_path.name}: page {page_index}/{total_pages} 完了 / {config['label']}")
    return records


def extract_native_pdf_page(file_path, page, page_index):
    records = []
    for word in page.get_text("words"):
        x0, y0, x1, y1, text, block_no, line_no, word_no = word[:8]
        r = base_record(file_path, "pdf")
        r.update({"sheet_or_page": page_index, "element_type": "word", "text": text,
                  "x0": round(x0, 2), "y0": round(y0, 2), "x1": round(x1, 2), "y1": round(y1, 2),
                  "note": f"block:{block_no} line:{line_no} word:{word_no}"})
        records.append(r)
    try:
        finder = page.find_tables()
        for ti, table in enumerate(finder.tables, start=1):
            for ri, row in enumerate(table.extract(), start=1):
                for ci, value in enumerate(row, start=1):
                    if value is not None and str(value).strip():
                        r = base_record(file_path, "pdf")
                        r.update({"sheet_or_page": page_index, "element_type": "table_cell", "text": str(value).strip(),
                                  "row": ri, "column": ci, "table_id": f"page:{page_index}/table:{ti}"})
                        records.append(r)
    except Exception:
        pass
    return records


def extract_ocr_pdf_page(file_path, page, page_index, ocr_engine, log_callback=None, ocr_mode="standard"):
    records = []
    log = log_callback or (lambda *_: None)
    config = OCR_MODES[ocr_mode]
    dpi = config["dpi"]
    pix = page.get_pixmap(dpi=dpi, alpha=False, colorspace=fitz.csGRAY)
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        tmp_path = Path(tmp.name)
    try:
        pix.save(str(tmp_path))
        log(f"  page {page_index}: {dpi}dpi・グレースケール画像化完了 / {config['label']}")
        results, api_name = run_paddle_ocr(ocr_engine, str(tmp_path))
        log(f"  page {page_index}: PaddleOCR API={api_name}")
        result_count = 0
        for result in results:
            parsed = parse_paddle_result(result)
            texts = parsed.get("rec_texts", [])
            scores = parsed.get("rec_scores", [])
            boxes = parsed.get("rec_boxes", [])
            log(f"  page {page_index}: OCR候補={len(texts)}")
            for i, text in enumerate(texts):
                text = str(text).strip()
                if not text:
                    continue
                score = scores[i] if i < len(scores) else ""
                box = boxes[i] if i < len(boxes) else None
                x0 = y0 = x1 = y1 = ""
                if box is not None:
                    try:
                        if len(box) == 4 and not isinstance(box[0], (list, tuple)):
                            x0, y0, x1, y1 = [float(v) for v in box]
                        else:
                            xs = [float(p[0]) for p in box]
                            ys = [float(p[1]) for p in box]
                            x0, x1 = min(xs), max(xs)
                            y0, y1 = min(ys), max(ys)
                        scale = 72 / dpi
                        x0 *= scale; y0 *= scale; x1 *= scale; y1 *= scale
                    except Exception:
                        x0 = y0 = x1 = y1 = ""
                r = base_record(file_path, "pdf")
                r.update({"sheet_or_page": page_index, "element_type": "ocr_text", "text": text,
                          "x0": round(x0, 2) if x0 != "" else "", "y0": round(y0, 2) if y0 != "" else "",
                          "x1": round(x1, 2) if x1 != "" else "", "y1": round(y1, 2) if y1 != "" else "",
                          "confidence": round(float(score), 4) if score != "" else "", "ocr_engine": "PaddleOCR",
                          "note": f"{config['label']} / {dpi}dpi grayscale / {config['det_model']} + {config['rec_model']} / lang=japan"})
                records.append(r)
                result_count += 1
        log(f"  page {page_index}: OCR取得={result_count}件")
    finally:
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass
    return records


def run_paddle_ocr(ocr_engine, image_path):
    """PaddleOCR 3.x の predict() と旧APIの ocr() の両方に対応する。"""
    if hasattr(ocr_engine, "predict"):
        try:
            return ocr_engine.predict(image_path), "predict"
        except Exception:
            # predict() が存在していても環境差で失敗する場合は旧APIも試す。
            pass

    if hasattr(ocr_engine, "ocr"):
        legacy = ocr_engine.ocr(image_path, cls=True)
        return legacy or [], "ocr(legacy)"

    raise RuntimeError("PaddleOCRに predict() / ocr() のどちらも見つかりません。")


def parse_paddle_result(result):
    # PaddleOCR 2.x 系: [box, (text, score)] の並びを吸収する。
    if isinstance(result, list):
        legacy_items = result
        if len(result) == 1 and isinstance(result[0], list):
            legacy_items = result[0]
        texts, scores, boxes = [], [], []
        legacy_ok = False
        for item in legacy_items:
            try:
                box, rec = item
                text, score = rec
                texts.append(text)
                scores.append(score)
                boxes.append(box)
                legacy_ok = True
            except Exception:
                legacy_ok = False
                break
        if legacy_ok and texts:
            return {"rec_texts": texts, "rec_scores": scores, "rec_boxes": boxes}

    candidates = []
    if isinstance(result, dict):
        candidates.append(result)
    try:
        value = result.json
        if callable(value):
            value = value()
        if isinstance(value, str):
            value = json.loads(value)
        if isinstance(value, dict):
            candidates.append(value)
    except Exception:
        pass
    for attr in ("to_dict", "dict"):
        try:
            fn = getattr(result, attr)
            value = fn() if callable(fn) else fn
            if isinstance(value, dict):
                candidates.append(value)
        except Exception:
            pass
    try:
        candidates.append({"rec_texts": result["rec_texts"], "rec_scores": result["rec_scores"], "rec_boxes": result["rec_boxes"]})
    except Exception:
        pass
    for data in candidates:
        if "res" in data and isinstance(data["res"], dict):
            data = data["res"]
        if "rec_texts" in data:
            return {"rec_texts": list(data.get("rec_texts", [])),
                    "rec_scores": list(data.get("rec_scores", [])),
                    "rec_boxes": list(data.get("rec_boxes", data.get("rec_polys", [])))}
    return {"rec_texts": [], "rec_scores": [], "rec_boxes": []}


FIELDS = ["source_file", "source_path", "source_type", "sheet_or_page", "element_type", "original_text", "text",
          "x0", "y0", "x1", "y1", "row", "column", "table_id", "confidence", "ocr_engine", "note"]




def display_ocr_confidence(value):
    """OCR信頼度が存在しない直接抽出データは「―」と表示する。"""
    if value is None or value == "":
        return "―"
    try:
        return f"{float(value):.2f}"
    except (TypeError, ValueError):
        return str(value)


def display_element_type(value):
    """内部の要素種別をユーザー向けの表示名へ変換する。"""
    value = str(value)
    if value == "word":
        return "文字列"
    return value

def export_csv_custom(records, fields, output_path):
    with output_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        for record in records:
            row = dict(record)
            if "element_type" in fields:
                row["element_type"] = display_element_type(record.get("element_type", ""))
            writer.writerow(row)


def _safe_excel_value(value):
    if value is None:
        return ""
    if isinstance(value, str):
        # XMLで禁止される制御文字を除去する。OCR文字列に混ざってもExcelを壊さない。
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    return value


def _write_excel_row_as_values(ws, row_index, values):
    for column_index, raw_value in enumerate(values, start=1):
        value = _safe_excel_value(raw_value)
        cell = ws.cell(row=row_index, column=column_index)
        cell.value = value
        # OCR結果が =SUM(...) や =□□ のように始まっても数式扱いさせない。
        if isinstance(value, str):
            cell.data_type = "s"


def export_xlsx_custom(records, fields, output_path, metadata=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "extracted_data"
    _write_excel_row_as_values(ws, 1, fields)
    for row_index, record in enumerate(records, start=2):
        row_values = []
        for field in fields:
            value = record.get(field, "")
            if field == "element_type":
                value = display_element_type(value)
            row_values.append(value)
        _write_excel_row_as_values(ws, row_index, row_values)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    if metadata:
        meta_ws = wb.create_sheet("data_quality")
        meta_ws.append(["item", "value"])
        for key, value in metadata.items():
            _write_excel_row_as_values(meta_ws, meta_ws.max_row + 1, [key, value])
        meta_ws.freeze_panes = "A2"
        meta_ws.column_dimensions["A"].width = 24
        meta_ws.column_dimensions["B"].width = 80

    wb.save(output_path)


def export_csv(records, output_path):
    with output_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS)
        writer.writeheader(); writer.writerows(records)


def export_xlsx(records, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "extracted_data"
    _write_excel_row_as_values(ws, 1, FIELDS)
    for row_index, record in enumerate(records, start=2):
        _write_excel_row_as_values(ws, row_index, [record.get(field, "") for field in FIELDS])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(output_path)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
