"""
太炅 Lotto Lab Ultimate v0.2

기능
- 역대 로또 Excel 불러오기
- 번호 빈도 / 페어 / 트리플 분석
- 사진 파일 목록 등록
- 번호 직접 입력 및 출현횟수 집계
- 역대 1등·2등 동일 조합 제외
- 조건 기반 추천 조합 생성
- 직접 만든 조합 검사
- 추천 결과 Excel 저장
"""

from __future__ import annotations

import math
import re
import sys
import traceback
from collections import Counter
from dataclasses import dataclass
from itertools import combinations
from pathlib import Path
from typing import Iterable

import pandas as pd
import cv2
import numpy as np
from rapidocr_onnxruntime import RapidOCR
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QFont
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFormLayout, QFrame,
    QGridLayout, QHBoxLayout, QLabel, QLineEdit, QListWidget, QMainWindow,
    QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QSpinBox,
    QStackedWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget
)

APP_NAME = "太炅 Lotto Lab Ultimate"
VERSION = "1.0.0"


@dataclass(frozen=True)
class Draw:
    round_no: int
    numbers: tuple[int, int, int, int, int, int]
    bonus: int | None


def parse_numbers(text: str) -> list[int]:
    values = [int(x) for x in re.findall(r"\d+", text)]
    invalid = [x for x in values if not 1 <= x <= 45]
    if invalid:
        raise ValueError(f"1~45 범위를 벗어난 번호: {invalid}")
    return values


class LottoAnalyzer:
    def __init__(self) -> None:
        self.draws: list[Draw] = []
        self.number_counts: Counter[int] = Counter()
        self.pair_counts: Counter[tuple[int, int]] = Counter()
        self.triple_counts: Counter[tuple[int, int, int]] = Counter()
        self.first_prize: set[tuple[int, ...]] = set()
        self.second_prize: set[tuple[int, ...]] = set()

    def load_excel(self, path: str | Path) -> None:
        xls = pd.ExcelFile(path)
        best: list[Draw] = []

        for sheet in xls.sheet_names:
            try:
                raw = pd.read_excel(path, sheet_name=sheet, header=None)
            except Exception:
                continue
            draws = self._parse_sheet(raw)
            if len(draws) > len(best):
                best = draws

        if not best:
            raise ValueError(
                "회차와 당첨번호 6개를 찾지 못했습니다. "
                "첫 행에 회차·당첨번호·보너스가 있는 파일을 사용하세요."
            )

        self.draws = sorted(best, key=lambda d: d.round_no)
        self._analyze()

    @staticmethod
    def _parse_sheet(df: pd.DataFrame) -> list[Draw]:
        """엑셀 시트에서 회차·당첨번호 6개·보너스를 안전하게 추출합니다."""
        if df.empty or df.shape[1] < 7:
            return []

        header_row = None
        for i in range(min(30, len(df))):
            texts = [str(v).strip().lower() for v in df.iloc[i].tolist()]
            if any("회차" in x or x == "round" for x in texts):
                header_row = i
                break

        if header_row is None:
            return []

        headers = [
            "" if pd.isna(v) else str(v).strip()
            for v in df.iloc[header_row].tolist()
        ]
        body = df.iloc[header_row + 1:].reset_index(drop=True)

        def find_index(predicate) -> int | None:
            for idx, header in enumerate(headers):
                if predicate(header):
                    return idx
            return None

        round_idx = find_index(
            lambda h: "회차" in h.lower() or h.lower() == "round"
        )
        bonus_idx = find_index(
            lambda h: "보너스" in h.lower() or "bonus" in h.lower()
        )

        if round_idx is None:
            return []

        order_words = ("첫번째", "두번째", "세번째", "네번째", "다섯번째", "여섯번째")
        number_indices: list[int] = []
        for word in order_words:
            idx = find_index(lambda h, w=word: w in h)
            if idx is not None:
                number_indices.append(idx)

        if len(number_indices) < 6:
            number_indices = []
            for n in range(1, 7):
                idx = find_index(
                    lambda h, n=n: bool(
                        re.search(rf"(번호|num|ball)\s*{n}$", h, re.I)
                    )
                )
                if idx is not None:
                    number_indices.append(idx)

        if len(number_indices) < 6:
            candidates: list[int] = []
            for idx in range(df.shape[1]):
                if idx in (round_idx, bonus_idx):
                    continue
                series = pd.to_numeric(body.iloc[:, idx], errors="coerce").dropna()
                if len(series) >= 10 and float(series.between(1, 45).mean()) >= 0.85:
                    candidates.append(idx)
            number_indices = candidates[:6]

        if len(number_indices) < 6:
            return []

        draws: list[Draw] = []
        for _, row in body.iterrows():
            try:
                round_no = int(float(row.iloc[round_idx]))
                nums = tuple(
                    sorted(int(float(row.iloc[idx])) for idx in number_indices[:6])
                )
            except (ValueError, TypeError, IndexError):
                continue

            if len(set(nums)) != 6 or not all(1 <= x <= 45 for x in nums):
                continue

            bonus = None
            if bonus_idx is not None:
                try:
                    bonus_value = row.iloc[bonus_idx]
                    if pd.notna(bonus_value):
                        candidate = int(float(bonus_value))
                        if 1 <= candidate <= 45:
                            bonus = candidate
                except (ValueError, TypeError, IndexError):
                    bonus = None

            draws.append(Draw(round_no, nums, bonus))

        return draws

    def _analyze(self) -> None:
        self.number_counts.clear()
        self.pair_counts.clear()
        self.triple_counts.clear()
        self.first_prize.clear()
        self.second_prize.clear()

        for draw in self.draws:
            self.number_counts.update(draw.numbers)
            self.pair_counts.update(combinations(draw.numbers, 2))
            self.triple_counts.update(combinations(draw.numbers, 3))
            self.first_prize.add(draw.numbers)

            # 2등 조합 = 본번호 5개 + 보너스번호
            if draw.bonus is not None:
                for five in combinations(draw.numbers, 5):
                    self.second_prize.add(tuple(sorted((*five, draw.bonus))))

    def check_combo(self, combo: tuple[int, ...]) -> dict:
        combo = tuple(sorted(combo))
        same_first = combo in self.first_prize
        same_second = combo in self.second_prize
        matches: list[tuple[int, int]] = []
        s = set(combo)
        for draw in self.draws:
            count = len(s.intersection(draw.numbers))
            if count >= 4:
                matches.append((draw.round_no, count))
        matches.sort(key=lambda x: (-x[1], -x[0]))
        return {"first": same_first, "second": same_second, "matches": matches}


class Recommender:
    def __init__(self, analyzer: LottoAnalyzer) -> None:
        self.a = analyzer

    @staticmethod
    def consecutive_pairs(combo: tuple[int, ...]) -> int:
        return sum(1 for a, b in zip(combo, combo[1:]) if b - a == 1)

    def score(self, combo: tuple[int, ...], source_weights: Counter[int]) -> float:
        # 각각의 점수는 순위용 통계 점수이며 당첨확률이 아님
        source = sum(source_weights[n] for n in combo) * 12.0
        freq = sum(self.a.number_counts[n] for n in combo) / max(1, len(self.a.draws))
        pair = sum(self.a.pair_counts[p] for p in combinations(combo, 2)) / 15.0
        triple = sum(self.a.triple_counts[t] for t in combinations(combo, 3)) / 20.0
        return source + freq + pair * 0.8 + triple * 1.2

    def generate(
        self,
        source_weights: Counter[int],
        count: int,
        sum_min: int,
        sum_max: int,
        allow_consecutive: bool,
    ) -> list[tuple[float, tuple[int, ...]]]:
        pool = sorted(source_weights)
        if len(pool) < 6:
            raise ValueError("고유 번호가 최소 6개 필요합니다.")

        candidates = []
        for combo in combinations(pool, 6):
            total = sum(combo)
            if not sum_min <= total <= sum_max:
                continue

            odd = sum(n % 2 for n in combo)
            if odd not in (2, 3, 4):
                continue

            high = sum(n >= 23 for n in combo)
            if high not in (2, 3, 4):
                continue

            if not allow_consecutive and self.consecutive_pairs(combo) > 0:
                continue
            if self.consecutive_pairs(combo) > 2:
                continue

            if combo in self.a.first_prize or combo in self.a.second_prize:
                continue

            candidates.append((self.score(combo, source_weights), combo))

        candidates.sort(key=lambda x: (-x[0], x[1]))
        return candidates[:count]


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.analyzer = LottoAnalyzer()
        self.photo_paths: list[str] = []
        self.recommendations: list[tuple[float, tuple[int, ...]]] = []
        self.ocr_engine = None

        self.setWindowTitle(f"{APP_NAME} v{VERSION}")
        self.resize(1320, 850)
        self.setMinimumSize(1100, 700)

        self.stack = QStackedWidget()
        self.pages = [
            self.make_dashboard(),
            self.make_source_page(),
            self.make_stats_page(),
            self.make_recommend_page(),
            self.make_checker_page(),
        ]
        for p in self.pages:
            self.stack.addWidget(p)

        root = QWidget()
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(self.make_sidebar())
        layout.addWidget(self.stack, 1)
        self.setCentralWidget(root)

        self.make_menu()
        self.apply_theme()
        self.statusBar().showMessage("역대 로또 Excel 파일을 불러오세요.")

    def make_sidebar(self) -> QWidget:
        box = QFrame()
        box.setObjectName("sidebar")
        box.setFixedWidth(245)
        lay = QVBoxLayout(box)
        lay.setContentsMargins(18, 22, 18, 18)

        logo = QLabel("太炅")
        logo.setObjectName("logo")
        logo.setAlignment(Qt.AlignCenter)
        sub = QLabel("Lotto Lab Ultimate")
        sub.setObjectName("subtitle")
        sub.setAlignment(Qt.AlignCenter)
        lay.addWidget(logo)
        lay.addWidget(sub)
        lay.addSpacing(20)

        names = ["대시보드", "사진·번호 입력", "통계 분석", "추천 조합", "조합 검사"]
        for i, name in enumerate(names):
            b = QPushButton(name)
            b.clicked.connect(lambda checked=False, idx=i: self.stack.setCurrentIndex(idx))
            lay.addWidget(b)

        lay.addStretch()
        b = QPushButton("역대 Excel 불러오기")
        b.setObjectName("primary")
        b.clicked.connect(self.open_excel)
        lay.addWidget(b)
        return box

    def make_menu(self) -> None:
        menu = self.menuBar().addMenu("파일")
        open_action = QAction("역대 Excel 불러오기", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_excel)
        menu.addAction(open_action)

        export_action = QAction("추천 결과 Excel 저장", self)
        export_action.triggered.connect(self.export_results)
        menu.addAction(export_action)

    def make_dashboard(self) -> QWidget:
        p = QWidget()
        lay = QVBoxLayout(p)
        title = QLabel("대시보드")
        title.setObjectName("pageTitle")
        lay.addWidget(title)

        self.dashboard_info = QLabel(
            "역대 로또 당첨번호 Excel을 불러오면 분석이 시작됩니다.\n\n"
            "사진을 추가하면 OCR로 숫자를 자동 인식하고, 오른쪽 입력란에 자동으로 추가합니다."
        )
        self.dashboard_info.setObjectName("card")
        self.dashboard_info.setAlignment(Qt.AlignCenter)
        self.dashboard_info.setMinimumHeight(230)
        lay.addWidget(self.dashboard_info)

        self.progress = QProgressBar()
        lay.addWidget(self.progress)
        lay.addStretch()
        return p

    def make_source_page(self) -> QWidget:
        p = QWidget()
        lay = QVBoxLayout(p)
        title = QLabel("사진·번호 입력")
        title.setObjectName("pageTitle")
        lay.addWidget(title)

        grid = QGridLayout()
        left = QFrame()
        left.setObjectName("card")
        ll = QVBoxLayout(left)
        ll.addWidget(QLabel("사진 파일 등록"))
        self.photo_list = QListWidget()
        ll.addWidget(self.photo_list)
        row = QHBoxLayout()
        add = QPushButton("사진 추가·자동 인식")
        add.clicked.connect(self.add_photos)
        delete = QPushButton("선택 삭제")
        delete.clicked.connect(self.delete_photo)
        rerun = QPushButton("선택 사진 다시 인식")
        rerun.clicked.connect(self.rerun_selected_photo_ocr)
        row.addWidget(add)
        row.addWidget(delete)
        ll.addLayout(row)
        ll.addWidget(rerun)

        right = QFrame()
        right.setObjectName("card")
        rl = QVBoxLayout(right)
        rl.addWidget(QLabel(
            "사진에서 인식된 번호가 자동으로 들어옵니다.\n"
            "같은 번호가 반복되면 출현횟수 가중치로 반영됩니다."
        ))
        self.source_input = QPlainTextEdit()
        self.source_input.setPlaceholderText(
            "예:\n16 29 42 12 13\n"
            "2 6 8 9 15 18 22 28 30 34 35 37"
        )
        rl.addWidget(self.source_input)
        analyze = QPushButton("입력 번호 집계")
        analyze.setObjectName("primary")
        analyze.clicked.connect(self.update_source_counts)
        rl.addWidget(analyze)
        self.source_summary = QLabel("입력 대기")
        self.source_summary.setWordWrap(True)
        rl.addWidget(self.source_summary)

        grid.addWidget(left, 0, 0)
        grid.addWidget(right, 0, 1)
        lay.addLayout(grid)
        return p

    def make_stats_page(self) -> QWidget:
        p = QWidget()
        lay = QVBoxLayout(p)
        title = QLabel("통계 분석")
        title.setObjectName("pageTitle")
        lay.addWidget(title)

        self.stats_type = QComboBox()
        self.stats_type.addItems(["번호 빈도", "페어 상위 100", "트리플 상위 100"])
        self.stats_type.currentIndexChanged.connect(self.refresh_stats_table)
        lay.addWidget(self.stats_type)

        self.stats_table = QTableWidget(0, 3)
        self.stats_table.setHorizontalHeaderLabels(["순위", "번호/조합", "출현 횟수"])
        self.stats_table.setEditTriggers(QTableWidget.NoEditTriggers)
        lay.addWidget(self.stats_table)
        return p

    def make_recommend_page(self) -> QWidget:
        p = QWidget()
        lay = QVBoxLayout(p)
        title = QLabel("추천 조합")
        title.setObjectName("pageTitle")
        lay.addWidget(title)

        opts = QFrame()
        opts.setObjectName("card")
        form = QFormLayout(opts)

        self.rec_count = QSpinBox()
        self.rec_count.setRange(1, 100)
        self.rec_count.setValue(20)

        self.sum_min = QSpinBox()
        self.sum_min.setRange(21, 255)
        self.sum_min.setValue(100)
        self.sum_max = QSpinBox()
        self.sum_max.setRange(21, 255)
        self.sum_max.setValue(180)

        self.allow_consecutive = QCheckBox("연속번호 허용")
        self.allow_consecutive.setChecked(True)

        form.addRow("추천 개수", self.rec_count)
        form.addRow("번호 합계 최소", self.sum_min)
        form.addRow("번호 합계 최대", self.sum_max)
        form.addRow("", self.allow_consecutive)

        run = QPushButton("추천 조합 생성")
        run.setObjectName("primary")
        run.clicked.connect(self.generate_recommendations)
        form.addRow("", run)
        lay.addWidget(opts)

        self.rec_table = QTableWidget(0, 6)
        self.rec_table.setHorizontalHeaderLabels(
            ["순위", "추천 조합", "분석 점수", "합계", "홀짝", "고저"]
        )
        self.rec_table.setEditTriggers(QTableWidget.NoEditTriggers)
        lay.addWidget(self.rec_table)

        export = QPushButton("추천 결과 Excel 저장")
        export.clicked.connect(self.export_results)
        lay.addWidget(export)
        return p

    def make_checker_page(self) -> QWidget:
        p = QWidget()
        lay = QVBoxLayout(p)
        title = QLabel("조합 검사")
        title.setObjectName("pageTitle")
        lay.addWidget(title)

        self.check_input = QLineEdit()
        self.check_input.setPlaceholderText("예: 12 16 22 29 34 42")
        lay.addWidget(self.check_input)

        btn = QPushButton("검사")
        btn.setObjectName("primary")
        btn.clicked.connect(self.check_combo)
        lay.addWidget(btn)

        self.check_result = QPlainTextEdit()
        self.check_result.setReadOnly(True)
        lay.addWidget(self.check_result)
        return p

    def open_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "역대 로또 당첨번호 Excel 선택", "",
            "Excel (*.xlsx *.xls)"
        )
        if not path:
            return
        try:
            self.progress.setValue(15)
            QApplication.processEvents()
            self.analyzer.load_excel(path)
            self.progress.setValue(100)
            latest = self.analyzer.draws[-1].round_no
            self.dashboard_info.setText(
                f"파일: {Path(path).name}\n"
                f"분석 회차: {len(self.analyzer.draws):,}회\n"
                f"최신 회차: {latest}회\n"
                f"1등 조합: {len(self.analyzer.first_prize):,}개\n"
                f"2등 성립 조합: {len(self.analyzer.second_prize):,}개\n\n"
                "번호·페어·트리플 분석 완료"
            )
            self.refresh_stats_table()
            self.statusBar().showMessage("Excel 분석 완료")
        except Exception as e:
            self.progress.setValue(0)
            QMessageBox.critical(self, "불러오기 오류", f"{e}\n\n{traceback.format_exc(limit=2)}")

    def _get_ocr_engine(self):
        if self.ocr_engine is None:
            self.statusBar().showMessage("OCR 엔진을 처음 준비하는 중입니다...")
            QApplication.processEvents()
            try:
                self.ocr_engine = RapidOCR()
            except Exception as exc:
                raise RuntimeError(
                    "OCR 엔진 초기화에 실패했습니다. "
                    "EXE에 OCR 모델·설정 파일이 포함되지 않았을 수 있습니다.\n"
                    f"상세 오류: {exc}"
                ) from exc
        return self.ocr_engine

    @staticmethod
    def _numbers_from_ocr_text(texts: list[str]) -> list[int]:
        numbers: list[int] = []
        for text in texts:
            # 시간(예: 6:08), 날짜, 페이지(1/6) 등은 최대한 제외
            if re.search(r"\d{1,2}:\d{2}", text):
                text = re.sub(r"\d{1,2}:\d{2}", " ", text)
            if re.search(r"\d{4}[./-]\d{1,2}[./-]\d{1,2}", text):
                text = re.sub(r"\d{4}[./-]\d{1,2}[./-]\d{1,2}", " ", text)
            text = re.sub(r"\b\d+\s*/\s*\d+\b", " ", text)
            text = re.sub(r"\b(96|100)\b", " ", text)  # 배터리 표시 오인식 방지

            for token in re.findall(r"(?<!\d)\d{1,2}(?!\d)", text):
                value = int(token)
                if 1 <= value <= 45:
                    numbers.append(value)
        return numbers

    def extract_numbers_from_photo(self, path: str) -> list[int]:
        # Windows의 한글·공백·긴 경로에서도 안전하게 사진을 읽습니다.
        try:
            file_bytes = np.fromfile(path, dtype=np.uint8)
            image = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
        except Exception as exc:
            raise ValueError(f"사진 파일을 열 수 없습니다: {exc}") from exc

        if image is None:
            raise ValueError(
                "사진 파일을 열 수 없습니다. 파일이 손상되었거나 지원하지 않는 형식일 수 있습니다."
            )

        # 휴대폰 상태표시줄과 하단 내비게이션 영역을 제외해 오인식을 줄임
        height, width = image.shape[:2]
        top = int(height * 0.09)
        bottom = int(height * 0.93)
        cropped = image[top:bottom, 0:width]

        # 글자 대비 강화
        gray = cv2.cvtColor(cropped, cv2.COLOR_BGR2GRAY)
        gray = cv2.convertScaleAbs(gray, alpha=1.35, beta=8)

        engine = self._get_ocr_engine()
        try:
            result, _ = engine(gray)
        except Exception as exc:
            raise ValueError(f"OCR 처리 중 오류가 발생했습니다: {exc}") from exc

        texts: list[str] = []
        if result:
            for item in result:
                if len(item) >= 2:
                    texts.append(str(item[1]))

        return self._numbers_from_ocr_text(texts)

    def append_ocr_numbers(self, numbers: list[int]) -> None:
        if not numbers:
            return
        current = self.source_input.toPlainText().rstrip()
        line = " ".join(map(str, numbers))
        self.source_input.setPlainText((current + "\n" + line).strip())
        self.update_source_counts()

    def add_photos(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "번호 사진 선택", "",
            "Images (*.png *.jpg *.jpeg *.webp *.bmp)"
        )
        if not paths:
            return

        all_numbers: list[int] = []
        failed: list[str] = []

        for path in paths:
            if path not in self.photo_paths:
                self.photo_paths.append(path)
                self.photo_list.addItem(Path(path).name)

            try:
                self.statusBar().showMessage(f"OCR 인식 중: {Path(path).name}")
                QApplication.processEvents()
                all_numbers.extend(self.extract_numbers_from_photo(path))
            except Exception:
                failed.append(Path(path).name)

        if all_numbers:
            self.append_ocr_numbers(all_numbers)
            self.statusBar().showMessage(
                f"사진 {len(paths)}장 OCR 완료 — 숫자 {len(all_numbers)}개 인식"
            )
        else:
            self.statusBar().showMessage("사진에서 1~45 숫자를 찾지 못했습니다.")

        message = (
            f"사진 {len(paths)}장 처리 완료\n"
            f"인식된 숫자: {len(all_numbers)}개"
        )
        if failed:
            message += "\n인식 실패: " + ", ".join(failed)
        QMessageBox.information(self, "사진 OCR 결과", message)

    def rerun_selected_photo_ocr(self) -> None:
        row = self.photo_list.currentRow()
        if row < 0:
            QMessageBox.information(self, "사진 선택", "다시 인식할 사진을 선택하세요.")
            return
        path = self.photo_paths[row]
        try:
            numbers = self.extract_numbers_from_photo(path)
            if numbers:
                self.append_ocr_numbers(numbers)
                QMessageBox.information(
                    self,
                    "OCR 완료",
                    f"{Path(path).name}\n숫자 {len(numbers)}개를 입력란에 추가했습니다."
                )
            else:
                QMessageBox.information(
                    self, "OCR 결과", "사진에서 1~45 숫자를 찾지 못했습니다."
                )
        except Exception as exc:
            QMessageBox.warning(self, "OCR 오류", str(exc))

    def delete_photo(self) -> None:
        row = self.photo_list.currentRow()
        if row >= 0:
            self.photo_list.takeItem(row)
            self.photo_paths.pop(row)

    def source_weights(self) -> Counter[int]:
        nums = parse_numbers(self.source_input.toPlainText())
        return Counter(nums)

    def update_source_counts(self) -> None:
        try:
            counts = self.source_weights()
            if not counts:
                self.source_summary.setText("입력된 번호가 없습니다.")
                return
            ranked = sorted(counts.items(), key=lambda x: (-x[1], x[0]))
            self.source_summary.setText(
                f"고유 번호 {len(counts)}개 / 전체 입력 {sum(counts.values())}개\n" +
                " · ".join(f"{n}번 {c}회" for n, c in ranked)
            )
        except Exception as e:
            QMessageBox.warning(self, "번호 입력 오류", str(e))

    def refresh_stats_table(self) -> None:
        idx = self.stats_type.currentIndex()
        if idx == 0:
            items = sorted(
                ((n, self.analyzer.number_counts[n]) for n in range(1, 46)),
                key=lambda x: (-x[1], x[0])
            )
        elif idx == 1:
            items = self.analyzer.pair_counts.most_common(100)
        else:
            items = self.analyzer.triple_counts.most_common(100)

        self.stats_table.setRowCount(len(items))
        for r, (key, count) in enumerate(items, 1):
            text = str(key) if isinstance(key, int) else " · ".join(map(str, key))
            self.stats_table.setItem(r - 1, 0, QTableWidgetItem(str(r)))
            self.stats_table.setItem(r - 1, 1, QTableWidgetItem(text))
            self.stats_table.setItem(r - 1, 2, QTableWidgetItem(str(count)))
        self.stats_table.resizeColumnsToContents()

    def generate_recommendations(self) -> None:
        if not self.analyzer.draws:
            QMessageBox.warning(self, "데이터 없음", "먼저 역대 로또 Excel을 불러오세요.")
            return
        try:
            weights = self.source_weights()
            recommender = Recommender(self.analyzer)
            self.recommendations = recommender.generate(
                weights,
                self.rec_count.value(),
                self.sum_min.value(),
                self.sum_max.value(),
                self.allow_consecutive.isChecked(),
            )
            if not self.recommendations:
                QMessageBox.information(
                    self, "결과 없음",
                    "조건을 만족하는 조합이 없습니다. 합계 범위나 번호 입력을 조정하세요."
                )
                return

            self.rec_table.setRowCount(len(self.recommendations))
            for r, (score, combo) in enumerate(self.recommendations, 1):
                odd = sum(x % 2 for x in combo)
                high = sum(x >= 23 for x in combo)
                values = [
                    str(r), " · ".join(map(str, combo)), f"{score:.2f}",
                    str(sum(combo)), f"{odd}:{6-odd}", f"{high}:{6-high}"
                ]
                for c, value in enumerate(values):
                    self.rec_table.setItem(r - 1, c, QTableWidgetItem(value))
            self.rec_table.resizeColumnsToContents()
            self.statusBar().showMessage(f"추천 조합 {len(self.recommendations)}개 생성 완료")
        except Exception as e:
            QMessageBox.warning(self, "추천 오류", str(e))

    def check_combo(self) -> None:
        if not self.analyzer.draws:
            QMessageBox.warning(self, "데이터 없음", "먼저 역대 로또 Excel을 불러오세요.")
            return
        try:
            nums = sorted(set(parse_numbers(self.check_input.text())))
            if len(nums) != 6:
                raise ValueError("서로 다른 번호 6개를 입력하세요.")
            combo = tuple(nums)
            result = self.analyzer.check_combo(combo)
            lines = [
                f"검사 조합: {' · '.join(map(str, combo))}",
                f"역대 1등과 동일: {'예' if result['first'] else '아니오'}",
                f"역대 2등 성립 조합과 동일: {'예' if result['second'] else '아니오'}",
                "",
                "과거 본번호 일치 회차(4개 이상):"
            ]
            if result["matches"]:
                lines += [f"- {round_no}회: {count}개 일치" for round_no, count in result["matches"][:30]]
            else:
                lines.append("- 없음")
            self.check_result.setPlainText("\n".join(lines))
        except Exception as e:
            QMessageBox.warning(self, "검사 오류", str(e))

    def export_results(self) -> None:
        if not self.recommendations:
            QMessageBox.information(self, "저장할 결과 없음", "먼저 추천 조합을 생성하세요.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "추천 결과 저장", "Taegyeong_Lotto_추천결과.xlsx", "Excel (*.xlsx)"
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        rows = []
        for rank, (score, combo) in enumerate(self.recommendations, 1):
            rows.append({
                "순위": rank,
                "번호1": combo[0], "번호2": combo[1], "번호3": combo[2],
                "번호4": combo[3], "번호5": combo[4], "번호6": combo[5],
                "분석점수": round(score, 2),
                "합계": sum(combo),
                "역대1등동일": "아니오",
                "역대2등동일": "아니오",
            })
        pd.DataFrame(rows).to_excel(path, index=False)
        QMessageBox.information(self, "저장 완료", path)

    def apply_theme(self) -> None:
        self.setStyleSheet("""
        QMainWindow, QWidget {
            background:#111111; color:#F4F0E6;
            font-family:"Malgun Gothic"; font-size:14px;
        }
        #sidebar { background:#080808; border-right:1px solid #4A3A12; }
        #logo { color:#D4AF37; font-size:48px; font-weight:800; }
        #subtitle { color:#E8D9A7; font-weight:700; }
        #pageTitle { color:#D4AF37; font-size:28px; font-weight:800; padding:8px; }
        #card { background:#1A1A1A; border:1px solid #4A3A12;
                border-radius:12px; padding:16px; }
        QPushButton {
            background:#252525; color:#F4F0E6; border:1px solid #3A3A3A;
            border-radius:8px; padding:11px; text-align:left;
        }
        QPushButton:hover { border-color:#D4AF37; background:#302817; }
        #primary { background:#D4AF37; color:#111111; font-weight:800; text-align:center; }
        QPlainTextEdit, QLineEdit, QListWidget, QTableWidget, QSpinBox, QComboBox {
            background:#181818; color:#F4F0E6; border:1px solid #404040;
            border-radius:7px; padding:6px;
        }
        QHeaderView::section {
            background:#2A2416; color:#F0D980; padding:8px; border:0;
        }
        QProgressBar { background:#222; border:1px solid #444; border-radius:7px; text-align:center; }
        QProgressBar::chunk { background:#D4AF37; border-radius:6px; }
        """)


def main() -> int:
    app = QApplication(sys.argv)
    app.setFont(QFont("Malgun Gothic", 10))
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
