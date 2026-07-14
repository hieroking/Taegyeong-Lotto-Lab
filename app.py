"""
太炅 Lotto Lab Ultimate v0.2

기능
- 역대 로또 Excel 불러오기
- 번호 빈도 / 페어 / 트리플 분석
- 사진 파일 목록 등록
- 번호 직접 입력 및 출현횟수 집계
- 역대 1등·2등 동일 조합 제외
- 조건 기반 추천조합 생성
- 직접 만든 조합 검사
- 추천 결과 Excel 저장
"""

from __future__ import annotations

import math
import re
import sys
import json
import base64
import os
import subprocess
import traceback
from collections import Counter
from dataclasses import dataclass
from itertools import combinations
from pathlib import Path
from typing import Iterable

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QFont, QColor, QBrush
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFormLayout, QFrame,
    QGridLayout, QHBoxLayout, QLabel, QLineEdit, QListWidget, QMainWindow,
    QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QSpinBox,
    QStackedWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget
)

APP_NAME = "太炅 Lotto Lab Ultimate"
VERSION = "3.4.0"



WINDOWS_OCR_PS = '$ErrorActionPreference = "Stop"\n[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n\nfunction Await($AsyncOperation, [Type]$ResultType) {\n    $methods = [System.WindowsRuntimeSystemExtensions].GetMethods() |\n        Where-Object {\n            $_.Name -eq "AsTask" -and\n            $_.IsGenericMethod -and\n            $_.GetParameters().Count -eq 1\n        }\n    $method = $methods | Select-Object -First 1\n    if ($null -eq $method) {\n        throw "Windows Runtime AsTask 메서드를 찾지 못했습니다."\n    }\n    $generic = $method.MakeGenericMethod($ResultType)\n    $task = $generic.Invoke($null, @($AsyncOperation))\n    $task.Wait()\n    return $task.Result\n}\n\ntry {\n    Add-Type -AssemblyName System.Runtime.WindowsRuntime\n\n    $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]\n    $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]\n    $null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType=WindowsRuntime]\n\n    $imagePath = $env:LOTTO_OCR_IMAGE\n    if ([string]::IsNullOrWhiteSpace($imagePath)) {\n        throw "사진 경로가 전달되지 않았습니다."\n    }\n    if (!(Test-Path -LiteralPath $imagePath)) {\n        throw "사진 파일을 찾을 수 없습니다: $imagePath"\n    }\n\n    $fullPath = [System.IO.Path]::GetFullPath($imagePath)\n    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($fullPath)) ([Windows.Storage.StorageFile])\n    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])\n    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])\n    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])\n\n    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()\n    if ($null -eq $engine) {\n        throw "Windows OCR 엔진을 만들 수 없습니다. Windows 설정에서 한국어 OCR 언어 기능을 설치하세요."\n    }\n\n    $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])\n    $text = $result.Text\n\n    # 휴대폰 캡처에서 자주 섞이는 시간/날짜/페이지 표시 제거\n    $clean = [regex]::Replace($text, \'\\b\\d{1,2}:\\d{2}\\b\', \' \')\n    $clean = [regex]::Replace($clean, \'\\b\\d{4}[./-]\\d{1,2}[./-]\\d{1,2}\\b\', \' \')\n    $clean = [regex]::Replace($clean, \'\\b\\d+\\s*/\\s*\\d+\\b\', \' \')\n\n    $numbers = @()\n    foreach ($m in [regex]::Matches($clean, \'(?<!\\d)\\d{1,2}(?!\\d)\')) {\n        $n = [int]$m.Value\n        if ($n -ge 1 -and $n -le 45) {\n            $numbers += $n\n        }\n    }\n\n    @{ ok = $true; text = $text; numbers = $numbers } |\n        ConvertTo-Json -Compress -Depth 4\n    exit 0\n}\ncatch {\n    @{ ok = $false; error = $_.Exception.Message; numbers = @() } |\n        ConvertTo-Json -Compress -Depth 4\n    exit 1\n}'


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
        self.recent_number_counts: Counter[int] = Counter()
        self.recent_pair_counts: Counter[tuple[int, int]] = Counter()
        self.recent_triple_counts: Counter[tuple[int, int, int]] = Counter()
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
        self.recent_number_counts.clear()
        self.recent_pair_counts.clear()
        self.recent_triple_counts.clear()
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

        # 최근패턴은 최신 100회를 기준으로 계산
        for draw in self.draws[-100:]:
            self.recent_number_counts.update(draw.numbers)
            self.recent_pair_counts.update(combinations(draw.numbers, 2))
            self.recent_triple_counts.update(combinations(draw.numbers, 3))

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
    """입력빈도·동반수·트리플·최근패턴을 자동 종합해 순위를 계산합니다."""

    CATEGORY_NAMES = {
        "추천조합": "composite",
        "나온횟수": "input",
        "동반수": "pair",
        "트리플": "triple",
        "최근패턴": "recent",
        "자체추천": "self",
    }

    def __init__(self, analyzer: LottoAnalyzer) -> None:
        self.a = analyzer
        self.max_pair = max(self.a.pair_counts.values(), default=1)
        self.max_triple = max(self.a.triple_counts.values(), default=1)
        self.max_recent_number = max(self.a.recent_number_counts.values(), default=1)
        self.max_recent_pair = max(self.a.recent_pair_counts.values(), default=1)

    @staticmethod
    def consecutive_pairs(combo: tuple[int, ...]) -> int:
        return sum(1 for a, b in zip(combo, combo[1:]) if b - a == 1)

    def pair_details(self, combo: tuple[int, ...], top_n: int = 3):
        details = [
            (pair, self.a.pair_counts[pair])
            for pair in combinations(combo, 2)
        ]
        details.sort(key=lambda item: (-item[1], item[0]))
        return details[:top_n]

    @staticmethod
    def _normalize(value: float, maximum: float) -> float:
        if maximum <= 0:
            return 0.0
        return max(0.0, min(100.0, value / maximum * 100.0))

    def metrics(
        self,
        combo: tuple[int, ...],
        source_weights: Counter[int],
    ) -> dict[str, float]:
        max_input = max(source_weights.values(), default=1)
        input_raw = sum(source_weights[n] for n in combo)
        input_score = self._normalize(input_raw, max_input * 6)

        pair_values = [self.a.pair_counts[p] for p in combinations(combo, 2)]
        pair_score = self._normalize(
            sum(sorted(pair_values, reverse=True)[:5]),
            self.max_pair * 5,
        )

        triple_values = [self.a.triple_counts[t] for t in combinations(combo, 3)]
        triple_score = self._normalize(
            sum(sorted(triple_values, reverse=True)[:5]),
            self.max_triple * 5,
        )

        recent_number = sum(self.a.recent_number_counts[n] for n in combo)
        recent_pair = sum(self.a.recent_pair_counts[p] for p in combinations(combo, 2))
        recent_score = (
            self._normalize(recent_number, self.max_recent_number * 6) * 0.55
            + self._normalize(recent_pair, self.max_recent_pair * 15) * 0.45
        )

        odd = sum(n % 2 for n in combo)
        high = sum(n >= 23 for n in combo)
        total = sum(combo)
        structure = 100.0
        structure -= abs(odd - 3) * 12
        structure -= abs(high - 3) * 10
        if total < 100:
            structure -= (100 - total) * 0.8
        elif total > 180:
            structure -= (total - 180) * 0.8
        structure -= max(0, self.consecutive_pairs(combo) - 1) * 15
        structure = max(0.0, min(100.0, structure))

        # 자동 종합 기준: 사용자가 따로 가중치를 조절하지 않아도 됨
        composite = (
            input_score * 0.30
            + pair_score * 0.25
            + triple_score * 0.20
            + recent_score * 0.15
            + structure * 0.10
        )

        return {
            "input": input_score,
            "pair": pair_score,
            "triple": triple_score,
            "recent": recent_score,
            "structure": structure,
            "composite": composite,
        }

    def category_score(self, metrics: dict[str, float], category: str) -> float:
        key = self.CATEGORY_NAMES.get(category, "composite")
        if key == "composite":
            return metrics["composite"]
        if key == "self":
            return metrics.get("self", metrics["composite"])
        # 항목별 순위는 해당 항목 70% + 자동 종합 30%
        return metrics[key] * 0.70 + metrics["composite"] * 0.30

    def self_number_scores(self) -> dict[int, float]:
        """역대 전체 데이터만으로 1~45 번호별 자체 점수를 계산합니다."""
        max_all = max(self.a.number_counts.values(), default=1)
        max_recent = max(self.a.recent_number_counts.values(), default=1)

        # 번호별 동반수 중심성: 다른 번호들과 같이 나온 횟수의 합
        pair_centrality = {}
        for number in range(1, 46):
            total = 0
            for other in range(1, 46):
                if number == other:
                    continue
                pair = tuple(sorted((number, other)))
                total += self.a.pair_counts[pair]
            pair_centrality[number] = total
        max_centrality = max(pair_centrality.values(), default=1)

        # 최신 출현 회차와 지연 정도
        latest_round = self.a.draws[-1].round_no if self.a.draws else 0
        last_seen = {n: 0 for n in range(1, 46)}
        for draw in self.a.draws:
            for n in draw.numbers:
                last_seen[n] = draw.round_no
        max_delay = max((latest_round - last_seen[n] for n in range(1, 46)), default=1)

        scores = {}
        for n in range(1, 46):
            all_score = self.a.number_counts[n] / max_all * 100
            recent_score = self.a.recent_number_counts[n] / max_recent * 100
            central_score = pair_centrality[n] / max_centrality * 100
            delay = latest_round - last_seen[n]
            delay_score = delay / max(1, max_delay) * 100

            scores[n] = (
                all_score * 0.35
                + recent_score * 0.30
                + central_score * 0.20
                + delay_score * 0.15
            )
        return scores

    def generate_self(
        self,
        count: int = 100,
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        """사진·직접입력 없이 역대 전체 당첨번호만으로 추천합니다."""
        number_scores = self.self_number_scores()

        # 45개 전체 조합(약 814만 개)은 느리므로 자체 점수 상위 28개를 우선 탐색합니다.
        # 구간 다양성을 위해 각 10번대 구간의 강한 번호도 포함합니다.
        ranked = sorted(number_scores, key=lambda n: (-number_scores[n], n))
        pool = set(ranked[:24])
        for low, high in [(1, 10), (11, 20), (21, 30), (31, 40), (41, 45)]:
            zone = [n for n in ranked if low <= n <= high][:2]
            pool.update(zone)

        # 최대 28개로 제한
        pool = sorted(pool, key=lambda n: (-number_scores[n], n))[:28]
        pool = sorted(pool)

        # metrics()의 나온횟수 점수를 역대 출현빈도로 계산하기 위한 가중치
        historical_weights = Counter({
            n: max(1, self.a.number_counts[n])
            for n in range(1, 46)
        })

        candidates = []
        for combo in combinations(pool, 6):
            total = sum(combo)
            if not 20 <= total <= 300:
                continue

            odd = sum(n % 2 for n in combo)
            high = sum(n >= 23 for n in combo)
            if odd not in (2, 3, 4) or high not in (2, 3, 4):
                continue
            if self.consecutive_pairs(combo) > 2:
                continue
            if combo in self.a.first_prize or combo in self.a.second_prize:
                continue

            metrics = self.metrics(combo, historical_weights)
            own_number_score = sum(number_scores[n] for n in combo) / 6.0

            # 자체추천 최종 점수: 번호 자체점수와 조합 통계의 자동 종합
            self_score = own_number_score * 0.45 + metrics["composite"] * 0.55
            metrics = dict(metrics)
            metrics["self"] = self_score
            candidates.append((self_score, combo, metrics))

        candidates.sort(key=lambda x: (-x[0], x[1]))
        return candidates[:count]

    def generate(
        self,
        source_weights: Counter[int],
        count: int,
        sum_min: int,
        sum_max: int,
        allow_consecutive: bool,
        category: str,
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        pool = sorted(source_weights)
        if len(pool) < 6:
            raise ValueError("고유 번호가 최소 6개 필요합니다.")

        candidates = []
        for combo in combinations(pool, 6):
            total = sum(combo)
            if not sum_min <= total <= sum_max:
                continue

            odd = sum(n % 2 for n in combo)
            high = sum(n >= 23 for n in combo)
            if odd not in (2, 3, 4) or high not in (2, 3, 4):
                continue
            if not allow_consecutive and self.consecutive_pairs(combo) > 0:
                continue
            if self.consecutive_pairs(combo) > 2:
                continue
            if combo in self.a.first_prize or combo in self.a.second_prize:
                continue

            metrics = self.metrics(combo, source_weights)
            score = self.category_score(metrics, category)
            candidates.append((score, combo, metrics))

        candidates.sort(key=lambda x: (-x[0], x[1]))
        return candidates[:count]


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.analyzer = LottoAnalyzer()
        self.photo_paths: list[str] = []
        self.recommendations: list[tuple[float, tuple[int, ...], dict[str, float]]] = []

        self.setWindowTitle(f"{APP_NAME} v{VERSION}")
        self.resize(1320, 850)
        self.setMinimumSize(1100, 700)

        self.stack = QStackedWidget()
        self.pages = [
            self.make_source_page(),
            self.make_recommend_page(),
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

        names = ["번호 입력", "추천조합", "나온횟수", "동반수", "트리플", "최근패턴", "자체추천"]
        for i, name in enumerate(names):
            b = QPushButton(name)
            if i == 0:
                b.clicked.connect(lambda checked=False: self.stack.setCurrentIndex(0))
            else:
                category_name = name
                b.clicked.connect(
                    lambda checked=False, cat=category_name: self.show_recommend_category(cat)
                )
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
            "사진을 추가하면 별도 OCR 파일 없이 Windows 내장 OCR로 번호를 자동 인식합니다."
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
            "사진 또는 메모에 나온 번호를 그대로 입력하세요.\n"
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

        title = QLabel("자동 추천 결과")
        title.setObjectName("pageTitle")
        lay.addWidget(title)

        guide = QLabel(
            "사진 또는 직접 입력한 번호를 기준으로 자동 계산합니다.\n"
            "추천 100조합 · 번호합계 20~300 · 역대 1등·2등 동일 조합 제외\n자체추천은 사진이나 직접입력 없이 역대 전체 당첨번호만으로 계산합니다."
        )
        guide.setObjectName("card")
        guide.setWordWrap(True)
        lay.addWidget(guide)

        self.rec_category = QComboBox()
        self.rec_category.addItems(
            ["추천조합", "나온횟수", "동반수", "트리플", "최근패턴", "자체추천"]
        )
        self.rec_category.currentTextChanged.connect(self.generate_recommendations)
        lay.addWidget(self.rec_category)

        self.rec_table = QTableWidget(0, 10)
        self.rec_table.setHorizontalHeaderLabels(
            [
                "순위", "추천조합", "카테고리 점수", "종합 점수",
                "나온횟수", "동반수", "트리플", "최근패턴",
                "같이 잘 나온 수", "합계"
            ]
        )
        self.rec_table.setEditTriggers(QTableWidget.NoEditTriggers)
        lay.addWidget(self.rec_table)

        export = QPushButton("현재 순위 Excel 저장")
        export.clicked.connect(self.export_results)
        lay.addWidget(export)
        return p

    def show_recommend_category(self, category: str) -> None:
        self.stack.setCurrentIndex(1)
        index = self.rec_category.findText(category)
        if index >= 0:
            self.rec_category.setCurrentIndex(index)
        else:
            self.generate_recommendations()

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
            self.statusBar().showMessage("Excel 분석 완료 — 사진 또는 번호를 입력하세요.")
            self.stack.setCurrentIndex(0)
        except Exception as e:
            self.progress.setValue(0)
            QMessageBox.critical(self, "불러오기 오류", f"{e}\n\n{traceback.format_exc(limit=2)}")

    def run_windows_ocr(self, image_path: str) -> list[int]:
        """외부 파일 없이 Windows 10/11 내장 OCR을 호출합니다."""
        if sys.platform != "win32":
            raise RuntimeError("사진 OCR은 Windows 10/11에서만 사용할 수 있습니다.")

        encoded = base64.b64encode(
            WINDOWS_OCR_PS.encode("utf-16le")
        ).decode("ascii")

        env = os.environ.copy()
        env["LOTTO_OCR_IMAGE"] = str(Path(image_path).resolve())

        command = [
            "powershell.exe",
            "-NoLogo",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy", "Bypass",
            "-EncodedCommand", encoded,
        ]

        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8-sig",
            errors="replace",
            timeout=120,
            env=env,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )

        stdout = completed.stdout.strip()
        stderr = completed.stderr.strip()

        # PowerShell이 JSON 앞에 공백/경고를 붙인 경우 마지막 JSON 객체를 찾음
        json_line = ""
        for line in reversed(stdout.splitlines()):
            line = line.strip()
            if line.startswith("{") and line.endswith("}"):
                json_line = line
                break

        if not json_line:
            detail = stderr or stdout or "Windows OCR에서 결과를 받지 못했습니다."
            raise RuntimeError(detail[:1000])

        try:
            payload = json.loads(json_line)
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                "Windows OCR 결과를 해석하지 못했습니다.\n"
                f"출력: {json_line[:500]}"
            ) from exc

        if not payload.get("ok"):
            raise RuntimeError(payload.get("error", "Windows OCR 처리 실패"))

        numbers = payload.get("numbers") or []
        return [int(n) for n in numbers if 1 <= int(n) <= 45]

    def append_ocr_numbers(self, numbers: list[int]) -> None:
        if not numbers:
            return
        current = self.source_input.toPlainText().rstrip()
        added = " ".join(map(str, numbers))
        self.source_input.setPlainText((current + "\n" + added).strip())
        self.update_source_counts()

    def add_photos(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "번호 사진 선택", "",
            "Images (*.png *.jpg *.jpeg *.webp *.bmp)"
        )
        if not paths:
            return

        all_numbers: list[int] = []
        failures: list[str] = []

        for path in paths:
            if path not in self.photo_paths:
                self.photo_paths.append(path)
                self.photo_list.addItem(Path(path).name)

            try:
                self.statusBar().showMessage(f"내장 OCR 인식 중: {Path(path).name}")
                QApplication.processEvents()
                all_numbers.extend(self.run_windows_ocr(path))
            except Exception as exc:
                failures.append(f"{Path(path).name}: {exc}")

        if all_numbers:
            self.append_ocr_numbers(all_numbers)
            self.statusBar().showMessage(
                f"사진 {len(paths)}장 처리 완료 — 숫자 {len(all_numbers)}개 인식"
            )
        else:
            self.statusBar().showMessage("사진에서 1~45 숫자를 찾지 못했습니다.")

        message = (
            f"사진 {len(paths)}장 처리 완료\n"
            f"인식된 숫자: {len(all_numbers)}개"
        )
        if failures:
            message += "\n\n일부 오류:\n" + "\n".join(failures[:5])
        QMessageBox.information(self, "사진 OCR 결과", message)

    def rerun_selected_photo_ocr(self) -> None:
        row = self.photo_list.currentRow()
        if row < 0:
            QMessageBox.information(self, "사진 선택", "다시 인식할 사진을 선택하세요.")
            return

        path = self.photo_paths[row]
        try:
            numbers = self.run_windows_ocr(path)
            if numbers:
                self.append_ocr_numbers(numbers)
                QMessageBox.information(
                    self,
                    "OCR 완료",
                    f"{Path(path).name}\n숫자 {len(numbers)}개를 입력란에 추가했습니다."
                )
            else:
                QMessageBox.information(
                    self,
                    "OCR 결과",
                    "사진에서 1~45 숫자를 찾지 못했습니다."
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
            if self.analyzer.draws and len(counts) >= 6:
                self.generate_recommendations()
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

    def generate_recommendations(self, *_args) -> None:
        if not self.analyzer.draws:
            return
        try:
            recommender = Recommender(self.analyzer)
            category = self.rec_category.currentText()

            if category == "자체추천":
                self.recommendations = recommender.generate_self(100)
            else:
                weights = self.source_weights()
                self.recommendations = recommender.generate(
                    weights,
                    100,
                    20,
                    300,
                    True,
                    category,
                )
            if not self.recommendations:
                QMessageBox.information(
                    self, "결과 없음",
                    "조건을 만족하는 조합이 없습니다. 합계 범위나 입력 번호를 조정하세요."
                )
                return

            key_map = Recommender.CATEGORY_NAMES
            selected_key = key_map.get(category, "composite")
            self.rec_table.setRowCount(len(self.recommendations))

            for r, (score, combo, metrics) in enumerate(self.recommendations, 1):
                pair_text = ", ".join(
                    f"{a}↔{b} {count}회"
                    for (a, b), count in recommender.pair_details(combo, 3)
                )
                values = [
                    str(r),
                    " · ".join(map(str, combo)),
                    f"{score:.1f}",
                    f"{metrics['composite']:.1f}",
                    f"{metrics['input']:.1f}",
                    f"{metrics['pair']:.1f}",
                    f"{metrics['triple']:.1f}",
                    f"{metrics['recent']:.1f}",
                    pair_text,
                    str(sum(combo)),
                ]
                for c, value in enumerate(values):
                    self.rec_table.setItem(r - 1, c, QTableWidgetItem(value))

                # 현재 선택한 카테고리의 점수 칸을 강조
                metric_column = {
                    "composite": 3, "input": 4, "pair": 5,
                    "triple": 6, "recent": 7, "self": 2,
                }[selected_key]
                metric_value = metrics.get(selected_key, score)
                item = self.rec_table.item(r - 1, metric_column)
                if metric_value >= 70:
                    color = QColor("#2E7D32")
                elif metric_value >= 50:
                    color = QColor("#9A7B16")
                else:
                    color = QColor("#5A3A3A")
                item.setBackground(QBrush(color))
                item.setForeground(QBrush(QColor("#FFFFFF")))

                combo_item = self.rec_table.item(r - 1, 1)
                if metric_value >= 70:
                    combo_item.setForeground(QBrush(QColor("#7CFF8A")))
                elif metric_value >= 50:
                    combo_item.setForeground(QBrush(QColor("#FFD95A")))

            self.rec_table.resizeColumnsToContents()
            self.statusBar().showMessage(
                f"{category} 기준 추천 {len(self.recommendations)}개 생성 완료"
            )
        except Exception as e:
            QMessageBox.warning(self, "추천 오류", str(e))


    def export_results(self) -> None:
        if not self.recommendations:
            QMessageBox.information(self, "저장할 결과 없음", "먼저 추천조합을 생성하세요.")
            return
        category = self.rec_category.currentText()
        path, _ = QFileDialog.getSaveFileName(
            self,
            "추천 결과 저장",
            f"Taegyeong_Lotto_{category}_추천결과.xlsx",
            "Excel (*.xlsx)",
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        recommender = Recommender(self.analyzer)
        rows = []
        for rank, (score, combo, metrics) in enumerate(self.recommendations, 1):
            rows.append({
                "카테고리": category,
                "순위": rank,
                "번호1": combo[0], "번호2": combo[1], "번호3": combo[2],
                "번호4": combo[3], "번호5": combo[4], "번호6": combo[5],
                "카테고리점수": round(score, 1),
                "종합점수": round(metrics["composite"], 1),
                "입력횟수점수": round(metrics["input"], 1),
                "동반수점수": round(metrics["pair"], 1),
                "트리플점수": round(metrics["triple"], 1),
                "최근패턴점수": round(metrics["recent"], 1),
                "자체추천점수": round(metrics.get("self", 0.0), 1),
                "같이잘나온수": ", ".join(
                    f"{a}-{b}({count}회)"
                    for (a, b), count in recommender.pair_details(combo, 3)
                ),
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
