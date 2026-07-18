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
import tempfile
import urllib.request
import urllib.parse
import shutil
from datetime import datetime
from collections import Counter, defaultdict
from dataclasses import dataclass
from itertools import combinations
from pathlib import Path
from typing import Iterable

import pandas as pd
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction, QFont, QColor, QBrush, QImage
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFormLayout, QFrame,
    QGridLayout, QHBoxLayout, QLabel, QLineEdit, QListWidget, QMainWindow,
    QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QSpinBox, QDialog,
    QStackedWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget
)

APP_NAME = "太炅 Lotto Lab Ultimate"
VERSION = "11.2.1-v27-v40-ocr-thread-fix"



WINDOWS_OCR_PS = '$ErrorActionPreference = "Stop"\n[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n\nfunction Await($AsyncOperation, [Type]$ResultType) {\n    $methods = [System.WindowsRuntimeSystemExtensions].GetMethods() |\n        Where-Object {\n            $_.Name -eq "AsTask" -and\n            $_.IsGenericMethod -and\n            $_.GetParameters().Count -eq 1\n        }\n    $method = $methods | Select-Object -First 1\n    if ($null -eq $method) {\n        throw "Windows Runtime AsTask 메서드를 찾지 못했습니다."\n    }\n    $generic = $method.MakeGenericMethod($ResultType)\n    $task = $generic.Invoke($null, @($AsyncOperation))\n    $task.Wait()\n    return $task.Result\n}\n\ntry {\n    Add-Type -AssemblyName System.Runtime.WindowsRuntime\n\n    $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]\n    $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]\n    $null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType=WindowsRuntime]\n\n    $imagePath = $env:LOTTO_OCR_IMAGE\n    if ([string]::IsNullOrWhiteSpace($imagePath)) {\n        throw "사진 경로가 전달되지 않았습니다."\n    }\n    if (!(Test-Path -LiteralPath $imagePath)) {\n        throw "사진 파일을 찾을 수 없습니다: $imagePath"\n    }\n\n    $fullPath = [System.IO.Path]::GetFullPath($imagePath)\n    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($fullPath)) ([Windows.Storage.StorageFile])\n    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])\n    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])\n    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])\n\n    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()\n    if ($null -eq $engine) {\n        throw "Windows OCR 엔진을 만들 수 없습니다. Windows 설정에서 한국어 OCR 언어 기능을 설치하세요."\n    }\n\n    $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])\n    $text = $result.Text\n\n    # 휴대폰 캡처에서 자주 섞이는 시간/날짜/페이지 표시 제거\n    $clean = [regex]::Replace($text, \'\\b\\d{1,2}:\\d{2}\\b\', \' \')\n    $clean = [regex]::Replace($clean, \'\\b\\d{4}[./-]\\d{1,2}[./-]\\d{1,2}\\b\', \' \')\n    $clean = [regex]::Replace($clean, \'\\b\\d+\\s*/\\s*\\d+\\b\', \' \')\n\n    $numbers = @()\n    foreach ($m in [regex]::Matches($clean, \'(?<!\\d)\\d{1,2}(?!\\d)\')) {\n        $n = [int]$m.Value\n        if ($n -ge 1 -and $n -le 45) {\n            $numbers += $n\n        }\n    }\n\n    @{ ok = $true; text = $text; numbers = $numbers } |\n        ConvertTo-Json -Compress -Depth 4\n    exit 0\n}\ncatch {\n    @{ ok = $false; error = $_.Exception.Message; numbers = @() } |\n        ConvertTo-Json -Compress -Depth 4\n    exit 1\n}'


class OCRWorker(QThread):
    progress = Signal(int, int, str)
    completed = Signal(list, list)

    def __init__(self, paths: list[str], ocr_callable, parent=None) -> None:
        super().__init__(parent)
        self.paths = list(paths)
        self.ocr_callable = ocr_callable

    def run(self) -> None:
        all_numbers: list[int] = []
        failures: list[str] = []
        total = len(self.paths)
        for index, path in enumerate(self.paths, 1):
            self.progress.emit(index, total, Path(path).name)
            try:
                all_numbers.extend(self.ocr_callable(path))
            except Exception as exc:
                failures.append(f"{Path(path).name}: {exc}")
        self.completed.emit(all_numbers, failures)


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



class TKPerformanceEngine:
    """1~1218회 워크포워드 검증으로 선택한 성과 중심 번호·조합 엔진."""

    FEATURE_NAMES = ['최근10회', '최근30회', '최근100회', '최근300회', '전체빈도', '미출현간격', '직전이월수', '2회전재등장', '끝수흐름', '직전번호동반수', '인접연속수', '간격수흐름']
    OPTIMIZED_WEIGHTS = [0.00378073, 0.116455048, 0.104777671, 0.073486328, 0.005267944, 0.221622601, 0.030370023, 0.076836631, 0.257287055, 0.041604862, 0.022926405, 0.045584697]
    OPTIMIZATION_RESULT = {'tested_settings': 30000, 'feature_names': ['최근10회', '최근30회', '최근100회', '최근300회', '전체빈도', '미출현간격', '직전이월수', '2회전재등장', '끝수흐름', '직전번호동반수', '인접연속수', '간격수흐름'], 'weights': {'최근10회': 0.003781, '최근30회': 0.116455, '최근100회': 0.104778, '최근300회': 0.073486, '전체빈도': 0.005268, '미출현간격': 0.221623, '직전이월수': 0.03037, '2회전재등장': 0.076837, '끝수흐름': 0.257287, '직전번호동반수': 0.041605, '인접연속수': 0.022926, '간격수흐름': 0.045585}, 'train': {'average_top15_hits': 2.088, 'three_plus_rate': 0.316, 'four_plus_rate': 0.088, 'max_hits': 5}, 'validation': {'average_top15_hits': 2.0797, 'three_plus_rate': 0.3116, 'four_plus_rate': 0.1014, 'max_hits': 6}, 'holdout': {'round_start': 1081, 'round_end': 1218, 'average_top15_hits': 2.1232, 'three_plus_rate': 0.3188, 'four_plus_rate': 0.1449, 'max_hits': 5, 'random_expected_hits': 2.0}, 'elapsed_seconds': 24.08}

    @staticmethod
    def normalize(values):
        values = list(map(float, values))
        low, high = min(values), max(values)
        if high <= low:
            return [0.0] * len(values)
        return [(value - low) / (high - low) for value in values]

    @classmethod
    def number_scores(cls, draws):
        if len(draws) < 30:
            raise ValueError("성과최적추천은 최소 30회 이상의 데이터가 필요합니다.")
        history = [tuple(draw.numbers) for draw in draws]
        flat = lambda rows: [n for row in rows for n in row]
        feature_columns = []

        for window in (10, 30, 100, 300):
            counts = Counter(flat(history[-window:]))
            feature_columns.append(cls.normalize([counts[n] for n in range(1, 46)]))

        all_counts = Counter(flat(history))
        feature_columns.append(cls.normalize([all_counts[n] for n in range(1, 46)]))

        last_seen = {n: -1 for n in range(1, 46)}
        for index, row in enumerate(history):
            for number in row:
                last_seen[number] = index
        gaps = [len(history) - 1 - last_seen[n] for n in range(1, 46)]
        feature_columns.append(cls.normalize(gaps))

        last = set(history[-1])
        previous = set(history[-2])
        feature_columns.append([1.0 if n in last else 0.0 for n in range(1, 46)])
        feature_columns.append([
            1.0 if n in previous and n not in last else 0.0
            for n in range(1, 46)
        ])

        ending_counts = Counter(n % 10 for n in flat(history[-30:]))
        feature_columns.append(
            cls.normalize([ending_counts[n % 10] for n in range(1, 46)])
        )

        partner = Counter()
        for row in history[-100:]:
            row_set = set(row)
            overlap = len(row_set & last)
            if overlap:
                for number in row_set - last:
                    partner[number] += overlap
        feature_columns.append(cls.normalize([partner[n] for n in range(1, 46)]))

        adjacent = []
        for n in range(1, 46):
            adjacent.append(
                1.0 if n not in last and any(abs(n - x) == 1 for x in last) else 0.0
            )
        feature_columns.append(adjacent)

        gap_counts = Counter()
        for row in history[-30:]:
            for a, b in combinations(sorted(row), 2):
                if 1 <= b - a <= 15:
                    gap_counts[b - a] += 1
        common_gaps = [gap for gap, _ in gap_counts.most_common(5)]
        interval = [0.0] * 45
        for rank, gap in enumerate(common_gaps):
            value = 1.0 - rank * 0.15
            for source in last:
                for candidate in (source - gap, source + gap):
                    if 1 <= candidate <= 45 and candidate not in last:
                        interval[candidate - 1] = max(interval[candidate - 1], value)
        feature_columns.append(interval)

        scores = {}
        details = {}
        for number in range(1, 46):
            contributions = []
            for name, weight, column in zip(
                cls.FEATURE_NAMES, cls.OPTIMIZED_WEIGHTS, feature_columns
            ):
                value = column[number - 1]
                contributions.append((name, value * weight))
            scores[number] = sum(value for _, value in contributions) * 100.0
            details[number] = sorted(
                contributions, key=lambda item: (-item[1], item[0])
            )
        return scores, details

    @classmethod
    def generate(
        cls,
        analyzer,
        count=100,
        fixed_numbers=(),
        excluded_numbers=(),
        candidate_numbers=(),
    ):
        scores, details = cls.number_scores(analyzer.draws)
        fixed_set = set(fixed_numbers)
        excluded_set = set(excluded_numbers)
        candidate_set = set(candidate_numbers)

        ranked = [
            n for n, _ in sorted(scores.items(), key=lambda item: (-item[1], item[0]))
            if n not in excluded_set
        ]
        pool = []
        for n in list(fixed_set | candidate_set) + ranked:
            if n not in pool and n not in excluded_set:
                pool.append(n)
            if len(pool) >= 20:
                break
        pool = sorted(pool)

        # 번호 선택과 조합 배치를 분리:
        # 후보 TOP20 안에서 조합을 만들고 구조·분산·동반출현을 별도 평가합니다.
        raw = []
        recent_pair = analyzer.recent_pair_counts
        for combo in combinations(pool, 6):
            combo_set = set(combo)
            if fixed_set and not fixed_set.issubset(combo_set):
                continue
            if excluded_set & combo_set:
                continue
            if combo in analyzer.first_prize or combo in analyzer.second_prize:
                continue
            odd = sum(n % 2 for n in combo)
            high = sum(n >= 23 for n in combo)
            total = sum(combo)
            zones = [
                sum(lo <= n <= hi for n in combo)
                for lo, hi in ((1, 10), (11, 20), (21, 30), (31, 40), (41, 45))
            ]
            if odd not in (2, 3, 4) or high not in (2, 3, 4):
                continue
            if not 95 <= total <= 185:
                continue
            if max(zones) > 3:
                continue

            number_score = sum(scores[n] for n in combo) / 6.0
            pair_score = sum(
                recent_pair[tuple(sorted(pair))]
                for pair in combinations(combo, 2)
            ) / 15.0
            candidate_bonus = len(combo_set & candidate_set) * 2.5
            consecutive = sum(b - a == 1 for a, b in zip(combo, combo[1:]))
            structure = 100.0
            structure -= abs(odd - 3) * 5.0
            structure -= abs(high - 3) * 4.0
            structure -= abs(total - 140) * 0.12
            structure -= max(0, consecutive - 1) * 8.0
            final_score = number_score * 0.72 + min(100.0, pair_score * 10) * 0.13 + structure * 0.15 + candidate_bonus

            top_reasons = []
            for number in combo:
                strongest = details[number][:2]
                top_reasons.append(
                    f"{number}번: " + ", ".join(name for name, _ in strongest)
                )

            metrics = {
                "performance": final_score,
                "composite": final_score,
                "input": 0.0,
                "pair": min(100.0, pair_score * 10),
                "triple": 0.0,
                "recent": number_score,
                "structure": structure,
                "pattern_votes": 0,
                "strategy": "성과최적엔진",
                "performance_reasons": top_reasons,
                "candidate_hits": len(combo_set & candidate_set),
                "candidate_bonus": candidate_bonus,
                "filter_mode": "성과최적화",
            }
            raw.append((final_score, combo, metrics))

        raw.sort(key=lambda row: (-row[0], row[1]))

        # 지나치게 비슷한 조합을 줄여 실제 추천 묶음의 포착 범위를 확대합니다.
        selected = []
        number_usage = Counter()
        for score, combo, metrics in raw:
            overlap5 = any(len(set(combo) & set(old_combo)) >= 5 for _, old_combo, _ in selected)
            usage_penalty = sum(number_usage[n] for n in combo) * 0.35
            adjusted = score - usage_penalty
            if overlap5 and len(selected) >= 10:
                continue
            selected.append((adjusted, combo, metrics))
            number_usage.update(combo)
            if len(selected) >= count:
                break
        selected.sort(key=lambda row: (-row[0], row[1]))
        return selected

class TKEngineAudit:
    RESULT = {'method': {'type': 'walk_forward_category_audit', 'round_start': 1189, 'round_end': 1218, 'rounds': 30, 'note': '공통 30,000개 최적화와 별도 항목별 검증', 'elapsed_seconds': 18.9}, 'results': [{'category': '최근패턴', 'mode': '기본', 'rounds': 30, 'top15_average_hits': 2.3, 'top15_3plus_rate': 0.3667, 'top15_4plus_rate': 0.1667, 'top15_max_hits': 5, 'combo100_average_best_hits': 2.1333, 'combo100_3plus_rate': 0.3333, 'combo100_4plus_rate': 0.0667, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 102.1016}, {'category': '특이패턴추천', 'mode': '이월수', 'rounds': 30, 'top15_average_hits': 2.3, 'top15_3plus_rate': 0.3667, 'top15_4plus_rate': 0.1, 'top15_max_hits': 5, 'combo100_average_best_hits': 2.2, 'combo100_3plus_rate': 0.3333, 'combo100_4plus_rate': 0.0667, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 100.5675}, {'category': '자체추천', 'mode': '장기미출현복귀', 'rounds': 30, 'top15_average_hits': 2.3333, 'top15_3plus_rate': 0.5, 'top15_4plus_rate': 0.0667, 'top15_max_hits': 4, 'combo100_average_best_hits': 2.2667, 'combo100_3plus_rate': 0.4667, 'combo100_4plus_rate': 0.0333, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 100.3664}, {'category': '자체추천', 'mode': '이월수', 'rounds': 30, 'top15_average_hits': 2.2333, 'top15_3plus_rate': 0.3333, 'top15_4plus_rate': 0.1667, 'top15_max_hits': 5, 'combo100_average_best_hits': 2.0667, 'combo100_3plus_rate': 0.3333, 'combo100_4plus_rate': 0.0333, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 98.4664}, {'category': '특이패턴추천', 'mode': '장기미출현복귀', 'rounds': 30, 'top15_average_hits': 2.2667, 'top15_3plus_rate': 0.4, 'top15_4plus_rate': 0.0667, 'top15_max_hits': 4, 'combo100_average_best_hits': 2.2, 'combo100_3plus_rate': 0.3667, 'combo100_4plus_rate': 0.0333, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 97.568}, {'category': '통합데이터추천', 'mode': '최근중심형', 'rounds': 30, 'top15_average_hits': 2.0667, 'top15_3plus_rate': 0.2333, 'top15_4plus_rate': 0.1667, 'top15_max_hits': 4, 'combo100_average_best_hits': 2.0, 'combo100_3plus_rate': 0.2333, 'combo100_4plus_rate': 0.1, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 94.3355}, {'category': '자체추천', 'mode': '끝수흐름', 'rounds': 30, 'top15_average_hits': 2.1, 'top15_3plus_rate': 0.4, 'top15_4plus_rate': 0.1333, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.9333, 'combo100_3plus_rate': 0.3667, 'combo100_4plus_rate': 0.0667, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 92.5326}, {'category': '성과최적추천', 'mode': '기본', 'rounds': 30, 'top15_average_hits': 2.0667, 'top15_3plus_rate': 0.3667, 'top15_4plus_rate': 0.1, 'top15_max_hits': 4, 'combo100_average_best_hits': 2.0, 'combo100_3plus_rate': 0.3667, 'combo100_4plus_rate': 0.0333, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 90.3335}, {'category': '통합데이터추천', 'mode': '균형형', 'rounds': 30, 'top15_average_hits': 2.0, 'top15_3plus_rate': 0.2667, 'top15_4plus_rate': 0.1333, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.9, 'combo100_3plus_rate': 0.2, 'combo100_4plus_rate': 0.1, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 89.9655}, {'category': '특이패턴추천', 'mode': '끝수흐름', 'rounds': 30, 'top15_average_hits': 2.1, 'top15_3plus_rate': 0.4, 'top15_4plus_rate': 0.0667, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.9, 'combo100_3plus_rate': 0.3, 'combo100_4plus_rate': 0.0333, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 88.967}, {'category': '자체추천', 'mode': '자동종합', 'rounds': 30, 'top15_average_hits': 2.0333, 'top15_3plus_rate': 0.3667, 'top15_4plus_rate': 0.0667, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.9333, 'combo100_3plus_rate': 0.3, 'combo100_4plus_rate': 0.0333, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 87.3656}, {'category': '자체추천', 'mode': '단기강세', 'rounds': 30, 'top15_average_hits': 1.9333, 'top15_3plus_rate': 0.2333, 'top15_4plus_rate': 0.1, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.9, 'combo100_3plus_rate': 0.2333, 'combo100_4plus_rate': 0.0667, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 85.9665}, {'category': '자체추천', 'mode': '간격수흐름', 'rounds': 30, 'top15_average_hits': 1.9333, 'top15_3plus_rate': 0.2333, 'top15_4plus_rate': 0.0667, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.9333, 'combo100_3plus_rate': 0.2333, 'combo100_4plus_rate': 0.0667, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 85.2006}, {'category': '통합데이터추천', 'mode': '입력중심형', 'rounds': 30, 'top15_average_hits': 1.9333, 'top15_3plus_rate': 0.3, 'top15_4plus_rate': 0.0667, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.9, 'combo100_3plus_rate': 0.2667, 'combo100_4plus_rate': 0.0667, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 84.801}, {'category': '통합데이터추천', 'mode': '장기형', 'rounds': 30, 'top15_average_hits': 2.0, 'top15_3plus_rate': 0.3333, 'top15_4plus_rate': 0.0667, 'top15_max_hits': 5, 'combo100_average_best_hits': 1.8333, 'combo100_3plus_rate': 0.2667, 'combo100_4plus_rate': 0.0, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 3, 'audit_score': 84.3341}, {'category': '나온횟수', 'mode': '기본', 'rounds': 30, 'top15_average_hits': 1.8333, 'top15_3plus_rate': 0.2333, 'top15_4plus_rate': 0.1333, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.8, 'combo100_3plus_rate': 0.2333, 'combo100_4plus_rate': 0.1, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 83.7645}, {'category': '동반수', 'mode': '기본', 'rounds': 30, 'top15_average_hits': 1.8333, 'top15_3plus_rate': 0.2333, 'top15_4plus_rate': 0.1333, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.8, 'combo100_3plus_rate': 0.2333, 'combo100_4plus_rate': 0.1, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 83.7645}, {'category': '트리플', 'mode': '기본', 'rounds': 30, 'top15_average_hits': 1.8333, 'top15_3plus_rate': 0.2333, 'top15_4plus_rate': 0.1333, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.8, 'combo100_3plus_rate': 0.2333, 'combo100_4plus_rate': 0.1, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 83.7645}, {'category': '특이패턴추천', 'mode': '단기강세', 'rounds': 30, 'top15_average_hits': 1.8667, 'top15_3plus_rate': 0.2, 'top15_4plus_rate': 0.1, 'top15_max_hits': 5, 'combo100_average_best_hits': 1.8, 'combo100_3plus_rate': 0.2, 'combo100_4plus_rate': 0.0667, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 82.7685}, {'category': '특이패턴추천', 'mode': '간격수흐름', 'rounds': 30, 'top15_average_hits': 1.9333, 'top15_3plus_rate': 0.3, 'top15_4plus_rate': 0.0333, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.8667, 'combo100_3plus_rate': 0.2667, 'combo100_4plus_rate': 0.0333, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 82.3974}, {'category': '통합데이터추천', 'mode': '자동최적형', 'rounds': 30, 'top15_average_hits': 1.8667, 'top15_3plus_rate': 0.2333, 'top15_4plus_rate': 0.0667, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.8333, 'combo100_3plus_rate': 0.2, 'combo100_4plus_rate': 0.0667, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 82.0026}, {'category': '자체추천', 'mode': '2회전재등장', 'rounds': 30, 'top15_average_hits': 1.8667, 'top15_3plus_rate': 0.3, 'top15_4plus_rate': 0.1333, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.7, 'combo100_3plus_rate': 0.2667, 'combo100_4plus_rate': 0.0, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 3, 'audit_score': 81.0665}, {'category': '특이패턴추천', 'mode': '동반수확장', 'rounds': 30, 'top15_average_hits': 1.9, 'top15_3plus_rate': 0.3333, 'top15_4plus_rate': 0.0333, 'top15_max_hits': 5, 'combo100_average_best_hits': 1.8333, 'combo100_3plus_rate': 0.3, 'combo100_4plus_rate': 0.0333, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 80.9976}, {'category': '특이패턴추천', 'mode': '연속수후보', 'rounds': 30, 'top15_average_hits': 1.8667, 'top15_3plus_rate': 0.3667, 'top15_4plus_rate': 0.0667, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.7333, 'combo100_3plus_rate': 0.2667, 'combo100_4plus_rate': 0.0667, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 80.8026}, {'category': '자체추천', 'mode': '동반수확장', 'rounds': 30, 'top15_average_hits': 1.8667, 'top15_3plus_rate': 0.3333, 'top15_4plus_rate': 0.0667, 'top15_max_hits': 5, 'combo100_average_best_hits': 1.7667, 'combo100_3plus_rate': 0.3, 'combo100_4plus_rate': 0.0333, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 80.3684}, {'category': '특이패턴추천', 'mode': '자동종합', 'rounds': 30, 'top15_average_hits': 1.8667, 'top15_3plus_rate': 0.3, 'top15_4plus_rate': 0.0333, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.8333, 'combo100_3plus_rate': 0.3, 'combo100_4plus_rate': 0.0, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 3, 'audit_score': 79.1661}, {'category': '자체추천', 'mode': '연속수후보', 'rounds': 30, 'top15_average_hits': 1.8667, 'top15_3plus_rate': 0.3333, 'top15_4plus_rate': 0.0333, 'top15_max_hits': 4, 'combo100_average_best_hits': 1.7333, 'combo100_3plus_rate': 0.2333, 'combo100_4plus_rate': 0.0, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 3, 'audit_score': 77.9661}, {'category': '특이패턴추천', 'mode': '2회전재등장', 'rounds': 30, 'top15_average_hits': 1.8667, 'top15_3plus_rate': 0.2667, 'top15_4plus_rate': 0.0333, 'top15_max_hits': 5, 'combo100_average_best_hits': 1.6333, 'combo100_3plus_rate': 0.2, 'combo100_4plus_rate': 0.0333, 'combo100_5plus_rate': 0.0, 'combo100_max_hits': 4, 'audit_score': 77.5986}]}
    @classmethod
    def rows(cls): return cls.RESULT["results"]
    @classmethod
    def find(cls, category, mode):
        return next((r for r in cls.rows() if r["category"]==category and r["mode"]==mode), None)


class TKEvolutionEngine:
    """1~1000회 초기학습 후 1001~최신 회차를 순차 성장한 추천엔진."""

    FEATURE_NAMES = ['최근10회', '최근30회', '최근100회', '최근300회', '전체빈도', '미출현간격', '직전이월수', '2회전재등장', '끝수흐름', '직전번호동반수', '인접연속수', '간격수흐름']
    WEIGHTS = [0.0035521908, 0.1188050095, 0.1042657332, 0.0658493329, 0.0061050046, 0.2127215793, 0.0178634787, 0.0870927178, 0.2616207932, 0.0442250048, 0.0207339052, 0.0571652502]
    RESULT = {'version': '7.2.0', 'feature_names': ['최근10회', '최근30회', '최근100회', '최근300회', '전체빈도', '미출현간격', '직전이월수', '2회전재등장', '끝수흐름', '직전번호동반수', '인접연속수', '간격수흐름'], 'initial_training': '1~1000회', 'growth_range': '1001~1132회', 'holdout_range': '1133~1218회', 'deployment_growth_range': '1001~1218회', 'growth_steps': 132, 'deployment_steps': 218, 'adoptions_growth': 50, 'adoptions_deployment': 93, 'base_weights': {'최근10회': 0.00378073, '최근30회': 0.11645505, '최근100회': 0.10477767, '최근300회': 0.07348633, '전체빈도': 0.00526794, '미출현간격': 0.2216226, '직전이월수': 0.03037002, '2회전재등장': 0.07683663, '끝수흐름': 0.25728706, '직전번호동반수': 0.04160486, '인접연속수': 0.02292641, '간격수흐름': 0.0455847}, 'validated_weights': {'최근10회': 0.00655552, '최근30회': 0.08927259, '최근100회': 0.11661006, '최근300회': 0.08721183, '전체빈도': 0.00533293, '미출현간격': 0.22944106, '직전이월수': 0.0274888, '2회전재등장': 0.09348839, '끝수흐름': 0.23039409, '직전번호동반수': 0.04071259, '인접연속수': 0.02020779, '간격수흐름': 0.05328436}, 'deployment_weights': {'최근10회': 0.00355219, '최근30회': 0.11880501, '최근100회': 0.10426573, '최근300회': 0.06584933, '전체빈도': 0.006105, '미출현간격': 0.21272158, '직전이월수': 0.01786348, '2회전재등장': 0.08709272, '끝수흐름': 0.26162079, '직전번호동반수': 0.044225, '인접연속수': 0.02073391, '간격수흐름': 0.05716525}, 'holdout_base': {'avg': 2.0, 'r3': 0.3023, 'r4': 0.1279, 'r5': 0.0, 'max': 4, 'reward': 1.5802}, 'holdout_evolved': {'avg': 2.0349, 'r3': 0.314, 'r4': 0.1047, 'r5': 0.0, 'max': 4, 'reward': 1.55}, 'honest_note': '배포 가중치는 1218회까지 성장시킨 값이며, 1133~1218회 성적은 별도의 검증용 성장엔진으로 측정했습니다.'}

    @staticmethod
    def generate(analyzer, count=100, fixed_numbers=(), excluded_numbers=(), candidate_numbers=()):
        original = TKPerformanceEngine.OPTIMIZED_WEIGHTS
        try:
            TKPerformanceEngine.OPTIMIZED_WEIGHTS = TKEvolutionEngine.WEIGHTS
            rows = TKPerformanceEngine.generate(
                analyzer, count,
                fixed_numbers=fixed_numbers,
                excluded_numbers=excluded_numbers,
                candidate_numbers=candidate_numbers,
            )
            converted = []
            for score, combo, metrics in rows:
                metrics = dict(metrics)
                metrics["strategy"] = "AI회차별성장"
                metrics["evolution"] = score
                metrics["evolution_version"] = TKEvolutionEngine.RESULT["version"]
                metrics["evolution_steps"] = TKEvolutionEngine.RESULT["deployment_steps"]
                converted.append((score, combo, metrics))
            return converted
        finally:
            TKPerformanceEngine.OPTIMIZED_WEIGHTS = original


class TKV8EvolutionLab:
    FEATURE_NAMES = ['최근10', '최근30', '최근100', '최근300', '전체빈도', '미출현간격', '직전이월', '2회전재등장', '끝수흐름', '직전동반', '인접연속', '간격수흐름', '구간모멘텀', '홀짝모멘텀', '고저모멘텀', '번호온도', '패턴수명', '회차DNA유사', '반대AI', '정보량다양성']
    CHAMPION_WEIGHTS = [0.155035018, 0.026252538, 0.006447226, 0.0799974314, 0.0097650026, 0.039092923, 0.0232807063, 0.0172895129, 0.0138285768, 0.0213653572, 0.0059461019, 0.0069676216, 0.0897739367, 0.0252749595, 0.1657477697, 0.0450142405, 0.0183192603, 0.0208426067, 0.0546646552, 0.1750945555]
    ENSEMBLE_WEIGHTS = [0.1569861055, 0.0266200349, 0.0060077749, 0.0819162478, 0.009365791, 0.0383353187, 0.0231737093, 0.0173019575, 0.0133264283, 0.0208167826, 0.0057224441, 0.0073539721, 0.0918137387, 0.0246671171, 0.162500609, 0.0421119059, 0.0185296636, 0.0208679931, 0.0531565873, 0.1794258185]
    RESULT = {'version': '8.0.0', 'data': {'initial_training': '1~1000회', 'growth': '1001~1132회', 'holdout': '1133~1218회', 'latest': 1218}, 'world_cup': {'population': 240, 'generations': 60, 'total_engine_evaluations': 14400, 'elite_count': 36}, 'champion': {'engine_index': 197, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0814, 'holdout_3plus': 0.3488, 'holdout_4plus': 0.0581, 'holdout_5plus': 0.0116, 'holdout_max': 5, 'weights': {'최근10': 0.15503502, '최근30': 0.02625254, '최근100': 0.00644723, '최근300': 0.07999743, '전체빈도': 0.009765, '미출현간격': 0.03909292, '직전이월': 0.02328071, '2회전재등장': 0.01728951, '끝수흐름': 0.01382858, '직전동반': 0.02136536, '인접연속': 0.0059461, '간격수흐름': 0.00696762, '구간모멘텀': 0.08977394, '홀짝모멘텀': 0.02527496, '고저모멘텀': 0.16574777, '번호온도': 0.04501424, '패턴수명': 0.01831926, '회차DNA유사': 0.02084261, '반대AI': 0.05466466, '정보량다양성': 0.17509456}}, 'ensemble': {'engines': [22, 19, 23, 18, 197], 'holdout_avg': 2.0698, 'holdout_4plus': 0.0581, 'holdout_5plus': 0.0116, 'holdout_max': 5, 'weights': {'최근10': 0.15698611, '최근30': 0.02662003, '최근100': 0.00600777, '최근300': 0.08191625, '전체빈도': 0.00936579, '미출현간격': 0.03833532, '직전이월': 0.02317371, '2회전재등장': 0.01730196, '끝수흐름': 0.01332643, '직전동반': 0.02081678, '인접연속': 0.00572244, '간격수흐름': 0.00735397, '구간모멘텀': 0.09181374, '홀짝모멘텀': 0.02466712, '고저모멘텀': 0.16250061, '번호온도': 0.04211191, '패턴수명': 0.01852966, '회차DNA유사': 0.02086799, '반대AI': 0.05315659, '정보량다양성': 0.17942582}}, 'failure_analysis': {'후보번호부족': 56, '중간성과': 29, '고성과': 1}, 'pattern_lifecycle': {'최근10': {'recent': 0.0229, 'long': 0.0171, 'life': '유지'}, '최근30': {'recent': 0.0106, 'long': 0.0113, 'life': '유지'}, '최근100': {'recent': 0.0266, 'long': -0.0042, 'life': '강세'}, '최근300': {'recent': -0.0065, 'long': -0.0, 'life': '유지'}, '전체빈도': {'recent': 0.0131, 'long': -0.0012, 'life': '유지'}, '미출현간격': {'recent': -0.0436, 'long': -0.0159, 'life': '약세'}, '직전이월': {'recent': -0.0, 'long': 0.0006, 'life': '유지'}, '2회전재등장': {'recent': 0.0205, 'long': 0.0108, 'life': '유지'}, '끝수흐름': {'recent': 0.0157, 'long': 0.0118, 'life': '유지'}, '직전동반': {'recent': 0.0118, 'long': -0.0006, 'life': '유지'}, '인접연속': {'recent': -0.0038, 'long': -0.0207, 'life': '유지'}, '간격수흐름': {'recent': 0.0046, 'long': -0.009, 'life': '유지'}, '구간모멘텀': {'recent': 0.0194, 'long': 0.0132, 'life': '유지'}, '홀짝모멘텀': {'recent': 0.0, 'long': 0.0001, 'life': '유지'}, '고저모멘텀': {'recent': -0.0013, 'long': -0.0014, 'life': '유지'}, '번호온도': {'recent': -0.0065, 'long': 0.0133, 'life': '유지'}, '패턴수명': {'recent': 0.0092, 'long': 0.0034, 'life': '유지'}, '회차DNA유사': {'recent': -0.0013, 'long': 0.0006, 'life': '유지'}, '반대AI': {'recent': 0.0065, 'long': -0.0133, 'life': '유지'}, '정보량다양성': {'recent': 0.0254, 'long': 0.0182, 'life': '유지'}}, 'next_top20': [15, 27, 31, 38, 10, 35, 25, 33, 19, 13, 17, 29, 3, 32, 24, 23, 44, 28, 37, 5], 'engine_results': [{'rank': 1, 'engine_id': 0, 'growth_avg': 2.3485, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0349, 'holdout_4plus': 0.0581, 'holdout_5plus': 0.0116, 'max_hits': 5}, {'rank': 2, 'engine_id': 156, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0349, 'holdout_4plus': 0.0465, 'holdout_5plus': 0.0116, 'max_hits': 5}, {'rank': 3, 'engine_id': 12, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0116, 'holdout_4plus': 0.0581, 'holdout_5plus': 0.0, 'max_hits': 4}, {'rank': 4, 'engine_id': 154, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0233, 'holdout_4plus': 0.0581, 'holdout_5plus': 0.0, 'max_hits': 4}, {'rank': 5, 'engine_id': 9, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0349, 'holdout_4plus': 0.0465, 'holdout_5plus': 0.0116, 'max_hits': 5}, {'rank': 6, 'engine_id': 10, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0233, 'holdout_4plus': 0.0465, 'holdout_5plus': 0.0116, 'max_hits': 5}, {'rank': 7, 'engine_id': 8, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0233, 'holdout_4plus': 0.0581, 'holdout_5plus': 0.0, 'max_hits': 4}, {'rank': 8, 'engine_id': 4, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0233, 'holdout_4plus': 0.0581, 'holdout_5plus': 0.0, 'max_hits': 4}, {'rank': 9, 'engine_id': 11, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0233, 'holdout_4plus': 0.0465, 'holdout_5plus': 0.0, 'max_hits': 4}, {'rank': 10, 'engine_id': 7, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0349, 'holdout_4plus': 0.0581, 'holdout_5plus': 0.0, 'max_hits': 4}, {'rank': 11, 'engine_id': 5, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0349, 'holdout_4plus': 0.0581, 'holdout_5plus': 0.0116, 'max_hits': 5}, {'rank': 12, 'engine_id': 6, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0, 'holdout_4plus': 0.0465, 'holdout_5plus': 0.0, 'max_hits': 4}, {'rank': 13, 'engine_id': 109, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0581, 'holdout_4plus': 0.0581, 'holdout_5plus': 0.0116, 'max_hits': 5}, {'rank': 14, 'engine_id': 3, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0116, 'holdout_4plus': 0.0465, 'holdout_5plus': 0.0116, 'max_hits': 5}, {'rank': 15, 'engine_id': 124, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0349, 'holdout_4plus': 0.0465, 'holdout_5plus': 0.0, 'max_hits': 4}, {'rank': 16, 'engine_id': 14, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0116, 'holdout_4plus': 0.0465, 'holdout_5plus': 0.0, 'max_hits': 4}, {'rank': 17, 'engine_id': 1, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0116, 'holdout_4plus': 0.0465, 'holdout_5plus': 0.0, 'max_hits': 4}, {'rank': 18, 'engine_id': 2, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 1.9884, 'holdout_4plus': 0.0465, 'holdout_5plus': 0.0, 'max_hits': 4}, {'rank': 19, 'engine_id': 13, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0233, 'holdout_4plus': 0.0465, 'holdout_5plus': 0.0116, 'max_hits': 5}, {'rank': 20, 'engine_id': 197, 'growth_avg': 2.3409, 'growth_4plus': 0.2045, 'growth_5plus': 0.0379, 'holdout_avg': 2.0814, 'holdout_4plus': 0.0581, 'holdout_5plus': 0.0116, 'max_hits': 5}], 'generation_log': [{'generation': 1, 'best_score': 2.660668, 'best_avg': 2.1364, 'best_4plus': 0.1364, 'best_5plus': 0.0303, 'population': 240}, {'generation': 2, 'best_score': 2.746395, 'best_avg': 2.1818, 'best_4plus': 0.1591, 'best_5plus': 0.0303, 'population': 240}, {'generation': 3, 'best_score': 2.816785, 'best_avg': 2.2273, 'best_4plus': 0.1894, 'best_5plus': 0.0227, 'population': 240}, {'generation': 4, 'best_score': 2.85274, 'best_avg': 2.2727, 'best_4plus': 0.1591, 'best_5plus': 0.0303, 'population': 240}, {'generation': 5, 'best_score': 2.85274, 'best_avg': 2.2727, 'best_4plus': 0.1591, 'best_5plus': 0.0303, 'population': 240}, {'generation': 6, 'best_score': 2.856907, 'best_avg': 2.2424, 'best_4plus': 0.1818, 'best_5plus': 0.0303, 'population': 240}, {'generation': 7, 'best_score': 2.879943, 'best_avg': 2.2424, 'best_4plus': 0.1742, 'best_5plus': 0.0379, 'population': 240}, {'generation': 8, 'best_score': 2.88629, 'best_avg': 2.2424, 'best_4plus': 0.1818, 'best_5plus': 0.0379, 'population': 240}, {'generation': 9, 'best_score': 2.953143, 'best_avg': 2.2879, 'best_4plus': 0.1894, 'best_5plus': 0.0379, 'population': 240}, {'generation': 10, 'best_score': 2.97557, 'best_avg': 2.3258, 'best_4plus': 0.1894, 'best_5plus': 0.0303, 'population': 240}, {'generation': 11, 'best_score': 2.97557, 'best_avg': 2.3258, 'best_4plus': 0.1894, 'best_5plus': 0.0303, 'population': 240}, {'generation': 12, 'best_score': 2.97557, 'best_avg': 2.3258, 'best_4plus': 0.1894, 'best_5plus': 0.0303, 'population': 240}, {'generation': 13, 'best_score': 2.978942, 'best_avg': 2.2879, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 14, 'best_score': 2.985379, 'best_avg': 2.3106, 'best_4plus': 0.197, 'best_5plus': 0.0379, 'population': 240}, {'generation': 15, 'best_score': 2.985379, 'best_avg': 2.3106, 'best_4plus': 0.197, 'best_5plus': 0.0379, 'population': 240}, {'generation': 16, 'best_score': 2.985379, 'best_avg': 2.3106, 'best_4plus': 0.197, 'best_5plus': 0.0379, 'population': 240}, {'generation': 17, 'best_score': 2.986592, 'best_avg': 2.3106, 'best_4plus': 0.1894, 'best_5plus': 0.0379, 'population': 240}, {'generation': 18, 'best_score': 2.986592, 'best_avg': 2.3106, 'best_4plus': 0.1894, 'best_5plus': 0.0379, 'population': 240}, {'generation': 19, 'best_score': 2.986592, 'best_avg': 2.3106, 'best_4plus': 0.1894, 'best_5plus': 0.0379, 'population': 240}, {'generation': 20, 'best_score': 2.996848, 'best_avg': 2.3182, 'best_4plus': 0.1894, 'best_5plus': 0.0379, 'population': 240}, {'generation': 21, 'best_score': 3.005939, 'best_avg': 2.3182, 'best_4plus': 0.197, 'best_5plus': 0.0379, 'population': 240}, {'generation': 22, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 23, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 24, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 25, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 26, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 27, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 28, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 29, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 30, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 31, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 32, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 33, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 34, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 35, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 36, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 37, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 38, 'best_score': 3.017089, 'best_avg': 2.3258, 'best_4plus': 0.1818, 'best_5plus': 0.0455, 'population': 240}, {'generation': 39, 'best_score': 3.026802, 'best_avg': 2.3258, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 40, 'best_score': 3.026802, 'best_avg': 2.3258, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 41, 'best_score': 3.026802, 'best_avg': 2.3258, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 42, 'best_score': 3.03224, 'best_avg': 2.3258, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 43, 'best_score': 3.03224, 'best_avg': 2.3258, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 44, 'best_score': 3.03224, 'best_avg': 2.3258, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 45, 'best_score': 3.040065, 'best_avg': 2.3333, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 46, 'best_score': 3.047893, 'best_avg': 2.3409, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 47, 'best_score': 3.047893, 'best_avg': 2.3409, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 48, 'best_score': 3.047893, 'best_avg': 2.3409, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 49, 'best_score': 3.047893, 'best_avg': 2.3409, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 50, 'best_score': 3.047893, 'best_avg': 2.3409, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 51, 'best_score': 3.047893, 'best_avg': 2.3409, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 52, 'best_score': 3.047893, 'best_avg': 2.3409, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 53, 'best_score': 3.047893, 'best_avg': 2.3409, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 54, 'best_score': 3.047893, 'best_avg': 2.3409, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 55, 'best_score': 3.047893, 'best_avg': 2.3409, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 56, 'best_score': 3.047893, 'best_avg': 2.3409, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 57, 'best_score': 3.054523, 'best_avg': 2.3485, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 58, 'best_score': 3.054523, 'best_avg': 2.3485, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 59, 'best_score': 3.054523, 'best_avg': 2.3485, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}, {'generation': 60, 'best_score': 3.054523, 'best_avg': 2.3485, 'best_4plus': 0.2045, 'best_5plus': 0.0379, 'population': 240}], 'lineage': [{'generation': 1, 'champion_engine': 103, 'score': 2.6606676449255477}, {'generation': 2, 'champion_engine': 86, 'score': 2.7463947723952424}, {'generation': 3, 'champion_engine': 214, 'score': 2.816785153702002}, {'generation': 4, 'champion_engine': 52, 'score': 2.8527396425991958}, {'generation': 5, 'champion_engine': 0, 'score': 2.8527396425991958}, {'generation': 6, 'champion_engine': 159, 'score': 2.8569070977154807}, {'generation': 7, 'champion_engine': 228, 'score': 2.8799432967213847}, {'generation': 8, 'champion_engine': 221, 'score': 2.8862897704491695}, {'generation': 9, 'champion_engine': 152, 'score': 2.953143137552268}, {'generation': 10, 'champion_engine': 67, 'score': 2.975569620998475}, {'generation': 11, 'champion_engine': 0, 'score': 2.975569620998475}, {'generation': 12, 'champion_engine': 0, 'score': 2.975569620998475}, {'generation': 13, 'champion_engine': 61, 'score': 2.978942219458177}, {'generation': 14, 'champion_engine': 143, 'score': 2.9853794245002825}, {'generation': 15, 'champion_engine': 0, 'score': 2.9853794245002825}, {'generation': 16, 'champion_engine': 0, 'score': 2.9853794245002825}, {'generation': 17, 'champion_engine': 70, 'score': 2.986591558180029}, {'generation': 18, 'champion_engine': 0, 'score': 2.986591558180029}, {'generation': 19, 'champion_engine': 0, 'score': 2.986591558180029}, {'generation': 20, 'champion_engine': 159, 'score': 2.996848059973763}, {'generation': 21, 'champion_engine': 102, 'score': 3.005938969064672}, {'generation': 22, 'champion_engine': 148, 'score': 3.017088564862318}, {'generation': 23, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 24, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 25, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 26, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 27, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 28, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 29, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 30, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 31, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 32, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 33, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 34, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 35, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 36, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 37, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 38, 'champion_engine': 0, 'score': 3.017088564862318}, {'generation': 39, 'champion_engine': 64, 'score': 3.0268021099741618}, {'generation': 40, 'champion_engine': 0, 'score': 3.0268021099741618}, {'generation': 41, 'champion_engine': 0, 'score': 3.0268021099741618}, {'generation': 42, 'champion_engine': 75, 'score': 3.032240080013833}, {'generation': 43, 'champion_engine': 0, 'score': 3.032240080013833}, {'generation': 44, 'champion_engine': 0, 'score': 3.032240080013833}, {'generation': 45, 'champion_engine': 90, 'score': 3.040064864378242}, {'generation': 46, 'champion_engine': 195, 'score': 3.0478929253964115}, {'generation': 47, 'champion_engine': 0, 'score': 3.0478929253964115}, {'generation': 48, 'champion_engine': 0, 'score': 3.0478929253964115}, {'generation': 49, 'champion_engine': 0, 'score': 3.0478929253964115}, {'generation': 50, 'champion_engine': 0, 'score': 3.0478929253964115}, {'generation': 51, 'champion_engine': 163, 'score': 3.0478929253964115}, {'generation': 52, 'champion_engine': 0, 'score': 3.0478929253964115}, {'generation': 53, 'champion_engine': 0, 'score': 3.0478929253964115}, {'generation': 54, 'champion_engine': 0, 'score': 3.0478929253964115}, {'generation': 55, 'champion_engine': 203, 'score': 3.0478929253964115}, {'generation': 56, 'champion_engine': 1, 'score': 3.0478929253964115}, {'generation': 57, 'champion_engine': 205, 'score': 3.0545229116407704}, {'generation': 58, 'champion_engine': 0, 'score': 3.0545229116407704}, {'generation': 59, 'champion_engine': 0, 'score': 3.0545229116407704}, {'generation': 60, 'champion_engine': 0, 'score': 3.0545229116407704}]}

    @classmethod
    def generate(cls, analyzer, count=100, fixed_numbers=(), excluded_numbers=(), candidate_numbers=()):
        original = TKPerformanceEngine.OPTIMIZED_WEIGHTS
        try:
            TKPerformanceEngine.OPTIMIZED_WEIGHTS = list(cls.ENSEMBLE_WEIGHTS[:12])
            rows = TKPerformanceEngine.generate(analyzer, count, fixed_numbers=fixed_numbers,
                excluded_numbers=excluded_numbers, candidate_numbers=candidate_numbers)
            out=[]
            for score,combo,metrics in rows:
                metrics=dict(metrics)
                metrics["v8_lab"]=True
                metrics["v8_engine"]="메타앙상블"
                metrics["v8_worldcup_evaluations"]=cls.RESULT["world_cup"]["total_engine_evaluations"]
                metrics["v8_holdout_avg"]=cls.RESULT["ensemble"]["holdout_avg"]
                metrics["v8_holdout_4plus"]=cls.RESULT["ensemble"]["holdout_4plus"]
                metrics["strategy"]="V8-AI연구소"
                metrics["evolution"]=score
                out.append((score,combo,metrics))
            return out
        finally:
            TKPerformanceEngine.OPTIMIZED_WEIGHTS = original


class TKCandidateLab:
    """후보번호 적중률을 먼저 검증하고 그 후보로 100조합을 생성하는 중간 연구엔진."""

    RESULT = {'version': '8.1.0-candidate-lab', 'method': {'initial_training': '1~1000회', 'growth': '1001~1132회', 'holdout': '1133~1218회', 'candidate_sizes': [10, 12, 15, 18, 20], 'selected_candidate_size': 15, 'engine': 'V8 메타앙상블 12피처 후보번호 엔진'}, 'growth_results': {'10': {'rounds': 132, 'average_hits': 1.3864, 'three_plus_rate': 0.1439, 'four_plus_rate': 0.0152, 'five_plus_rate': 0.0, 'six_hits_rate': 0.0, 'max_hits': 4, 'coverage_efficiency': 0.138636}, '12': {'rounds': 132, 'average_hits': 1.6894, 'three_plus_rate': 0.2273, 'four_plus_rate': 0.0606, 'five_plus_rate': 0.0076, 'six_hits_rate': 0.0, 'max_hits': 5, 'coverage_efficiency': 0.140783}, '15': {'rounds': 132, 'average_hits': 2.0909, 'three_plus_rate': 0.3712, 'four_plus_rate': 0.1439, 'five_plus_rate': 0.0076, 'six_hits_rate': 0.0, 'max_hits': 5, 'coverage_efficiency': 0.139394}, '18': {'rounds': 132, 'average_hits': 2.5076, 'three_plus_rate': 0.4848, 'four_plus_rate': 0.2348, 'five_plus_rate': 0.0379, 'six_hits_rate': 0.0, 'max_hits': 5, 'coverage_efficiency': 0.13931}, '20': {'rounds': 132, 'average_hits': 2.7652, 'three_plus_rate': 0.5833, 'four_plus_rate': 0.2955, 'five_plus_rate': 0.053, 'six_hits_rate': 0.0076, 'max_hits': 6, 'coverage_efficiency': 0.138258}}, 'holdout_results': {'10': {'rounds': 86, 'average_hits': 1.2326, 'three_plus_rate': 0.093, 'four_plus_rate': 0.0233, 'five_plus_rate': 0.0, 'six_hits_rate': 0.0, 'max_hits': 4, 'coverage_efficiency': 0.123256}, '12': {'rounds': 86, 'average_hits': 1.5116, 'three_plus_rate': 0.1744, 'four_plus_rate': 0.0233, 'five_plus_rate': 0.0, 'six_hits_rate': 0.0, 'max_hits': 4, 'coverage_efficiency': 0.125969}, '15': {'rounds': 86, 'average_hits': 1.907, 'three_plus_rate': 0.3256, 'four_plus_rate': 0.0581, 'five_plus_rate': 0.0116, 'six_hits_rate': 0.0, 'max_hits': 5, 'coverage_efficiency': 0.127132}, '18': {'rounds': 86, 'average_hits': 2.3023, 'three_plus_rate': 0.4419, 'four_plus_rate': 0.1395, 'five_plus_rate': 0.0465, 'six_hits_rate': 0.0, 'max_hits': 5, 'coverage_efficiency': 0.127907}, '20': {'rounds': 86, 'average_hits': 2.593, 'three_plus_rate': 0.5116, 'four_plus_rate': 0.2093, 'five_plus_rate': 0.0465, 'six_hits_rate': 0.0233, 'max_hits': 6, 'coverage_efficiency': 0.129651}}, 'latest_top15': [15, 27, 38, 31, 3, 35, 10, 45, 20, 19, 44, 30, 17, 33, 29], 'note': '후보번호 선택 능력과 100조합 배치 능력을 분리하기 위한 중간 연구판입니다.'}
    SELECTED_SIZE = 15
    WEIGHTS = [0.1569861055, 0.0266200349, 0.0060077749, 0.0819162478, 0.009365791, 0.0383353187, 0.0231737093, 0.0173019575, 0.0133264283, 0.0208167826, 0.0057224441, 0.0073539721]

    @classmethod
    def candidate_scores(cls, analyzer):
        original = TKPerformanceEngine.OPTIMIZED_WEIGHTS
        try:
            TKPerformanceEngine.OPTIMIZED_WEIGHTS = list(cls.WEIGHTS)
            return TKPerformanceEngine.number_scores(analyzer.draws)
        finally:
            TKPerformanceEngine.OPTIMIZED_WEIGHTS = original

    @classmethod
    def top_candidates(cls, analyzer, excluded_numbers=()):
        scores, details = cls.candidate_scores(analyzer)
        excluded = set(excluded_numbers)
        ranked = [
            n for n, _ in sorted(scores.items(), key=lambda item: (-item[1], item[0]))
            if n not in excluded
        ]
        return ranked[:cls.SELECTED_SIZE], scores, details

    @classmethod
    def generate(
        cls,
        analyzer,
        count=100,
        fixed_numbers=(),
        excluded_numbers=(),
        candidate_numbers=(),
    ):
        fixed_set = set(fixed_numbers)
        excluded_set = set(excluded_numbers)
        preferred_set = set(candidate_numbers)

        pool, scores, details = cls.top_candidates(analyzer, excluded_numbers)

        # 사용자 후보번호·필수번호가 빠지지 않도록 우선 포함
        merged = []
        for number in list(fixed_set | preferred_set) + pool:
            if number not in excluded_set and number not in merged:
                merged.append(number)
        pool = merged[:cls.SELECTED_SIZE]

        # 고정수 때문에 15개보다 적어졌다면 점수순으로 보충
        ranked_all = [
            n for n, _ in sorted(scores.items(), key=lambda item: (-item[1], item[0]))
            if n not in excluded_set
        ]
        for number in ranked_all:
            if number not in pool:
                pool.append(number)
            if len(pool) >= cls.SELECTED_SIZE:
                break
        pool = sorted(pool[:cls.SELECTED_SIZE])

        raw = []
        for combo in combinations(pool, 6):
            combo_set = set(combo)
            if fixed_set and not fixed_set.issubset(combo_set):
                continue
            if excluded_set & combo_set:
                continue
            if combo in analyzer.first_prize or combo in analyzer.second_prize:
                continue

            odd = sum(n % 2 for n in combo)
            high = sum(n >= 23 for n in combo)
            total = sum(combo)
            consecutive = sum(b - a == 1 for a, b in zip(combo, combo[1:]))

            if odd not in (2, 3, 4) or high not in (2, 3, 4):
                continue
            if not 90 <= total <= 190:
                continue
            if consecutive > 2:
                continue

            number_score = sum(scores[n] for n in combo) / 6.0
            pair_score = sum(
                analyzer.recent_pair_counts[tuple(sorted(pair))]
                for pair in combinations(combo, 2)
            ) / 15.0
            preferred_bonus = len(combo_set & preferred_set) * 2.5
            structure = 100.0
            structure -= abs(odd - 3) * 5.0
            structure -= abs(high - 3) * 4.0
            structure -= abs(total - 140) * 0.11
            structure -= max(0, consecutive - 1) * 7.0

            final_score = (
                number_score * 0.70
                + min(100.0, pair_score * 10) * 0.15
                + structure * 0.15
                + preferred_bonus
            )

            metrics = {
                "candidate_lab": True,
                "candidate_pool": pool,
                "candidate_pool_size": len(pool),
                "candidate_holdout_average": cls.RESULT["holdout_results"][str(cls.SELECTED_SIZE)]["average_hits"],
                "candidate_holdout_4plus": cls.RESULT["holdout_results"][str(cls.SELECTED_SIZE)]["four_plus_rate"],
                "candidate_holdout_5plus": cls.RESULT["holdout_results"][str(cls.SELECTED_SIZE)]["five_plus_rate"],
                "candidate_number_score": number_score,
                "candidate_pair_score": pair_score,
                "candidate_structure": structure,
                "strategy": "AI후보번호연구-15개",
                "evolution": final_score,
            }
            raw.append((final_score, combo, metrics))

        raw.sort(key=lambda row: (-row[0], row[1]))

        # 상위 100조합 분산선택
        selected = []
        for row in raw:
            combo = row[1]
            if len(selected) >= 10 and any(
                len(set(combo) & set(old[1])) >= 5 for old in selected
            ):
                continue
            selected.append(row)
            if len(selected) >= count:
                break
        if len(selected) < count:
            used = {row[1] for row in selected}
            for row in raw:
                if row[1] not in used:
                    selected.append(row)
                    used.add(row[1])
                if len(selected) >= count:
                    break
        return selected[:count]


class TKCandidateShrinkLab:
    RESULT = {'version': '8.2.0-candidate-shrink', 'protocol': {'initial_training': '1~1000회', 'training': '1001~1088회', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'challengers': 30019, 'features': 19, 'selection_rule': '학습 상위 후보를 조정검증에서 선택하고 보류구간은 마지막 평가에만 사용'}, 'goal': '후보 안에 당첨번호 6개를 유지하면서 후보 수를 단계적으로 줄이기', 'baseline': {'15': {'avg': 1.907, 'r3': 0.3256, 'r4': 0.0581, 'r5': 0.0116, 'r6': 0.0, 'max': 5}, '16': {'avg': 2.0349, 'r3': 0.3721, 'r4': 0.093, 'r5': 0.0116, 'r6': 0.0, 'max': 5}, '17': {'avg': 2.1744, 'r3': 0.407, 'r4': 0.1163, 'r5': 0.0233, 'r6': 0.0, 'max': 5}, '18': {'avg': 2.3023, 'r3': 0.4419, 'r4': 0.1395, 'r5': 0.0465, 'r6': 0.0, 'max': 5}, '19': {'avg': 2.4419, 'r3': 0.4767, 'r4': 0.1628, 'r5': 0.0465, 'r6': 0.0116, 'max': 6}, '20': {'avg': 2.593, 'r3': 0.5116, 'r4': 0.2093, 'r5': 0.0465, 'r6': 0.0233, 'max': 6}, '21': {'avg': 2.8023, 'r3': 0.5698, 'r4': 0.2907, 'r5': 0.0581, 'r6': 0.0233, 'max': 6}, '22': {'avg': 2.8721, 'r3': 0.593, 'r4': 0.3023, 'r5': 0.0814, 'r6': 0.0233, 'max': 6}, '25': {'avg': 3.186, 'r3': 0.7093, 'r4': 0.3953, 'r5': 0.1163, 'r6': 0.0233, 'max': 6}}, 'challenger': {'15': {'avg': 1.8488, 'r3': 0.2558, 'r4': 0.0465, 'r5': 0.0, 'r6': 0.0, 'max': 4}, '16': {'avg': 2.0465, 'r3': 0.3256, 'r4': 0.093, 'r5': 0.0, 'r6': 0.0, 'max': 4}, '17': {'avg': 2.186, 'r3': 0.3953, 'r4': 0.093, 'r5': 0.0, 'r6': 0.0, 'max': 4}, '18': {'avg': 2.3256, 'r3': 0.4535, 'r4': 0.1163, 'r5': 0.0, 'r6': 0.0, 'max': 4}, '19': {'avg': 2.4535, 'r3': 0.5116, 'r4': 0.1395, 'r5': 0.0116, 'r6': 0.0, 'max': 5}, '20': {'avg': 2.6163, 'r3': 0.5465, 'r4': 0.2209, 'r5': 0.0233, 'r6': 0.0, 'max': 5}, '21': {'avg': 2.7558, 'r3': 0.6163, 'r4': 0.2674, 'r5': 0.0233, 'r6': 0.0, 'max': 5}, '22': {'avg': 2.8837, 'r3': 0.6279, 'r4': 0.2791, 'r5': 0.093, 'r6': 0.0, 'max': 5}, '25': {'avg': 3.2558, 'r3': 0.7326, 'r4': 0.407, 'r5': 0.1744, 'r6': 0.0116, 'max': 6}}, 'decision': {'champion': '기존 V8 메타앙상블 후보엔진', 'challenger_status': '폐기', 'reason': '도전자는 20개 후보의 평균·3개·4개 성적은 소폭 개선했지만 6개 전부 포함 사례가 2회에서 0회로 감소했고, 5개 이상 성적도 하락했습니다.', 'smallest_candidate_with_six_all': 19, 'recommended_research_range': '19~22개에서 6개 전부 포함률을 보존한 뒤 점차 축소'}}


class TKDualCandidateLab:
    """기존 핵심후보를 보존하고 공격형 보조엔진 번호를 추가하는 이중 후보 연구엔진."""

    RESULT = {'version': '8.3.0-dual-candidate', 'protocol': {'initial_training': '1~1000회', 'training': '1001~1088회', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'search_count': 30000, 'feature_count': 9, 'holdout_used_for_selection': False}, 'design': {'core_engine': '기존 V8 메타앙상블', 'core_size': 20, 'support_engine': '9피처 고적중 보조엔진', 'support_add_count': 3, 'final_candidate_size': 23, 'rule': '기존 핵심 20개는 그대로 유지하고 보조엔진의 비중복 상위번호 3개만 추가'}, 'baseline_20': {'avg': 2.593, 'r4': 0.2093, 'r5': 0.0465, 'r6': 0.0233, 'max': 6}, 'baseline_23': {'avg': 2.9884, 'r4': 0.3372, 'r5': 0.1047, 'r6': 0.0233, 'max': 6}, 'hybrid_23': {'avg': 3.0233, 'r4': 0.3256, 'r5': 0.093, 'r6': 0.0349, 'max': 6}, 'change': {'average_hits': 0.0349, 'four_plus_pp': -1.16, 'five_plus_pp': -1.17, 'six_all_pp': 1.16, 'six_all_cases_before': 2, 'six_all_cases_after': 3}, 'decision': {'status': '조건부 채택', 'role': '공격형 보조 챔피언', 'stable_champion': '기존 V8 후보엔진 유지', 'reason': '6개 전부 포함은 2회에서 3회로 늘었지만 4개·5개 이상 지표는 소폭 하락했으므로 기존 엔진을 교체하지 않고 공격형 보조엔진으로 병행합니다.'}, 'latest': {'core20': [15, 27, 38, 31, 3, 35, 10, 45, 20, 19, 44, 30, 17, 33, 29, 7, 8, 23, 37, 13], 'support_add3': [34, 16, 6], 'hybrid23': [15, 27, 38, 31, 3, 35, 10, 45, 20, 19, 44, 30, 17, 33, 29, 7, 8, 23, 37, 13, 34, 16, 6]}, 'support_feature_weights': {'V8기준점수': 0.00632663, '자체추천점수': 0.02523131, '특이패턴평균': 0.45475509, '최근30빈도': 0.00259037, '최근100빈도': 0.09135935, '최근300빈도': 0.17758255, '전체빈도': 0.01369752, '최근동반중심성': 0.0206565, '미출현간격': 0.20780068}}
    SUPPORT_WEIGHTS = [0.00632663, 0.02523131, 0.45475509, 0.00259037, 0.09135935, 0.17758255, 0.01369752, 0.0206565, 0.20780068]

    @staticmethod
    def _normalize_map(values):
        raw = [float(values.get(n, 0.0)) for n in range(1, 46)]
        lo, hi = min(raw), max(raw)
        if hi <= lo:
            return {n: 0.0 for n in range(1, 46)}
        return {n: (raw[n-1] - lo) / (hi - lo) for n in range(1, 46)}

    @classmethod
    def _support_scores(cls, analyzer):
        recommender = Recommender(analyzer)

        original = TKPerformanceEngine.OPTIMIZED_WEIGHTS
        try:
            TKPerformanceEngine.OPTIMIZED_WEIGHTS = list(TKV8EvolutionLab.ENSEMBLE_WEIGHTS[:12])
            baseline, _ = TKPerformanceEngine.number_scores(analyzer.draws)
        finally:
            TKPerformanceEngine.OPTIMIZED_WEIGHTS = original

        self_scores = recommender.self_number_scores()
        pattern_scores, _ = recommender._pattern_number_scores(analyzer.draws)
        pattern_avg = {
            n: sum(pattern_scores[name][n] for name in recommender.PATTERN_NAMES)
               / len(recommender.PATTERN_NAMES)
            for n in range(1, 46)
        }

        f30 = Counter(n for draw in analyzer.draws[-30:] for n in draw.numbers)
        f100 = Counter(n for draw in analyzer.draws[-100:] for n in draw.numbers)
        f300 = Counter(n for draw in analyzer.draws[-300:] for n in draw.numbers)

        pair_center = {}
        for n in range(1, 46):
            pair_center[n] = sum(
                analyzer.recent_pair_counts[tuple(sorted((n, other)))]
                for other in range(1, 46) if other != n
            )

        last_seen = {n: -1 for n in range(1, 46)}
        for index, draw in enumerate(analyzer.draws):
            for number in draw.numbers:
                last_seen[number] = index
        gaps = {
            n: len(analyzer.draws) - 1 - last_seen[n]
            for n in range(1, 46)
        }

        maps = [
            cls._normalize_map(baseline),
            cls._normalize_map(self_scores),
            cls._normalize_map(pattern_avg),
            cls._normalize_map(f30),
            cls._normalize_map(f100),
            cls._normalize_map(f300),
            cls._normalize_map(analyzer.number_counts),
            cls._normalize_map(pair_center),
            cls._normalize_map(gaps),
        ]
        return {
            n: sum(cls.SUPPORT_WEIGHTS[i] * maps[i][n] for i in range(len(maps)))
            for n in range(1, 46)
        }

    @classmethod
    def candidate_pool(cls, analyzer, excluded_numbers=()):
        excluded = set(excluded_numbers)

        original = TKPerformanceEngine.OPTIMIZED_WEIGHTS
        try:
            TKPerformanceEngine.OPTIMIZED_WEIGHTS = list(TKV8EvolutionLab.ENSEMBLE_WEIGHTS[:12])
            baseline_scores, _ = TKPerformanceEngine.number_scores(analyzer.draws)
        finally:
            TKPerformanceEngine.OPTIMIZED_WEIGHTS = original

        baseline_rank = [
            n for n, _ in sorted(baseline_scores.items(), key=lambda item: (-item[1], item[0]))
            if n not in excluded
        ]
        support_scores = cls._support_scores(analyzer)
        support_rank = [
            n for n, _ in sorted(support_scores.items(), key=lambda item: (-item[1], item[0]))
            if n not in excluded
        ]

        pool = baseline_rank[:20]
        additions = []
        for number in support_rank:
            if number not in pool:
                pool.append(number)
                additions.append(number)
            if len(pool) >= 23:
                break
        return pool, additions, baseline_scores, support_scores

    @classmethod
    def generate(
        cls,
        analyzer,
        count=100,
        fixed_numbers=(),
        excluded_numbers=(),
        candidate_numbers=(),
    ):
        fixed_set = set(fixed_numbers)
        excluded_set = set(excluded_numbers)
        preferred_set = set(candidate_numbers)

        pool, additions, baseline_scores, support_scores = cls.candidate_pool(
            analyzer, excluded_numbers
        )

        merged = []
        for number in list(fixed_set | preferred_set) + pool:
            if number not in excluded_set and number not in merged:
                merged.append(number)

        # 사용자 지정 번호를 넣어도 최종 후보는 23개로 유지
        ranked_all = [
            n for n, _ in sorted(
                baseline_scores.items(),
                key=lambda item: (-item[1], item[0]),
            )
            if n not in excluded_set
        ]
        for number in ranked_all:
            if number not in merged:
                merged.append(number)
            if len(merged) >= 23:
                break
        pool = sorted(merged[:23])

        raw = []
        for combo in combinations(pool, 6):
            combo_set = set(combo)
            if fixed_set and not fixed_set.issubset(combo_set):
                continue
            if excluded_set & combo_set:
                continue
            if combo in analyzer.first_prize or combo in analyzer.second_prize:
                continue

            odd = sum(n % 2 for n in combo)
            high = sum(n >= 23 for n in combo)
            total = sum(combo)
            consecutive = sum(b - a == 1 for a, b in zip(combo, combo[1:]))

            if odd not in (2, 3, 4) or high not in (2, 3, 4):
                continue
            if not 85 <= total <= 200 or consecutive > 2:
                continue

            base_score = sum(baseline_scores[n] for n in combo) / 6
            support_score = sum(support_scores[n] for n in combo) / 6
            pair_score = sum(
                analyzer.recent_pair_counts[tuple(sorted(pair))]
                for pair in combinations(combo, 2)
            ) / 15
            preferred_bonus = len(combo_set & preferred_set) * 2.5

            structure = 100.0
            structure -= abs(odd - 3) * 4.0
            structure -= abs(high - 3) * 4.0
            structure -= abs(total - 140) * 0.10
            structure -= max(0, consecutive - 1) * 6.0

            final_score = (
                base_score * 0.52
                + support_score * 0.20
                + min(100.0, pair_score * 10) * 0.13
                + structure * 0.15
                + preferred_bonus
            )

            metrics = {
                "dual_candidate_lab": True,
                "candidate_pool": pool,
                "candidate_pool_size": len(pool),
                "support_additions": additions,
                "holdout_average": cls.RESULT["hybrid_23"]["avg"],
                "holdout_4plus": cls.RESULT["hybrid_23"]["r4"],
                "holdout_5plus": cls.RESULT["hybrid_23"]["r5"],
                "holdout_6all": cls.RESULT["hybrid_23"]["r6"],
                "strategy": "AI이중후보-공격형23",
                "evolution": final_score,
            }
            raw.append((final_score, combo, metrics))

        raw.sort(key=lambda row: (-row[0], row[1]))
        selected = []
        for row in raw:
            combo = row[1]
            if len(selected) >= 10 and any(
                len(set(combo) & set(old[1])) >= 5 for old in selected
            ):
                continue
            selected.append(row)
            if len(selected) >= count:
                break

        if len(selected) < count:
            used = {row[1] for row in selected}
            for row in raw:
                if row[1] not in used:
                    selected.append(row)
                    used.add(row[1])
                if len(selected) >= count:
                    break
        return selected[:count]


class TKCandidate22ShrinkLab:
    RESULT = {'version': '8.4.0-candidate22-shrink', 'protocol': {'initial_training': '1~1000회', 'training': '1001~1088회', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'search_space': 4059, 'selection': '학습 상위 100개 중 조정검증 최고 설정 선택', 'holdout_used_for_selection': False}, 'strategy': {'source_engine': 'v8.3 이중후보 23개', 'target_size': 22, 'baseline_weight': 0.475, 'support_weight': 0.525, 'support_addition_penalty': -0.2, 'core_preservation_bonus': 0.03, 'description': '핵심후보를 우선 보존하고, 보조추가 번호는 더 엄격하게 평가해 23개에서 22개로 축소'}, 'baseline_23': {'average_hits': 3.0233, 'three_plus_rate': 0.686, 'four_plus_rate': 0.3256, 'five_plus_rate': 0.093, 'six_all_rate': 0.0349, 'six_all_cases': 3, 'max_hits': 6}, 'challenger_22': {'average_hits': 2.8953, 'three_plus_rate': 0.6395, 'four_plus_rate': 0.314, 'five_plus_rate': 0.0814, 'six_all_rate': 0.0233, 'six_all_cases': 2, 'max_hits': 6}, 'change': {'average_hits': -0.128, 'three_plus_pp': -4.65, 'four_plus_pp': -1.16, 'five_plus_pp': -1.16, 'six_all_pp': -1.16, 'six_all_cases': -1}, 'decision': {'status': '공격형 조건부 채택', 'champion': 'v8.3 이중후보 23개 공격형 보조 챔피언 유지', 'reason': '후보를 1개 줄였지만 6개 전부 포함 사례가 3회에서 2회로 감소했고 평균·3개·4개·5개 지표도 모두 하락했습니다.', 'next_research': '23개 구조는 유지하고, 회차별로 제거 여부를 결정하는 동적 축소 방식 연구'}}


class TKDynamicCandidateSizeLab:
    RESULT = {'version': '8.5.0-dynamic-candidate-size', 'protocol': {'initial_training': '1~1000회', 'training': '1001~1088회', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'model': '표준화 + 균형 로지스틱 회귀', 'features': 11, 'threshold': 0.3, 'holdout_used_for_training_or_selection': False}, 'goal': '회차별로 후보 22개 또는 23개를 선택해 23개 엔진의 6개 전부 포함 성과를 보존하면서 평균 후보 수를 줄이기', 'baseline_23': {'average_candidate_size': 23.0, 'average_hits': 3.0233, 'three_plus_rate': 0.686, 'four_plus_rate': 0.3256, 'five_plus_rate': 0.093, 'six_all_rate': 0.0349, 'six_all_cases': 3, 'max_hits': 6}, 'static_22': {'average_candidate_size': 22.0, 'average_hits': 2.9186, 'three_plus_rate': 0.6395, 'four_plus_rate': 0.3023, 'five_plus_rate': 0.0814, 'six_all_rate': 0.0233, 'six_all_cases': 2, 'max_hits': 6}, 'dynamic_22_23': {'average_candidate_size': 22.6512, 'rounds_using_23': 56, 'rounds_using_22': 30, 'average_hits': 2.9651, 'three_plus_rate': 0.6744, 'four_plus_rate': 0.314, 'five_plus_rate': 0.0814, 'six_all_rate': 0.0233, 'six_all_cases': 2, 'max_hits': 6}, 'change_vs_baseline_23': {'average_candidate_size': -0.3488, 'average_hits': -0.0582, 'three_plus_pp': -1.16, 'four_plus_pp': -1.16, 'five_plus_pp': -1.16, 'six_all_pp': -1.16, 'six_all_cases': -1}, 'decision': {'status': '폐기', 'champion': 'v8.3 이중후보 23개 공격형 보조 챔피언 유지', 'reason': '평균 후보 수는 23개에서 22.65개로 0.35개 줄었지만 6개 전부 포함 사례가 3회에서 2회로 감소했고 평균·4개·5개 성적도 모두 하락했습니다.', 'next_research': '전체 후보 수를 바꾸지 말고 23개 경계번호의 교체 여부만 학습하는 방식'}}


class TKBoundarySwapLab:
    RESULT = {'version': '8.6.0-boundary-swap', 'protocol': {'initial_training': '1~1000회', 'training': '1001~1088회', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'model': '표준화 + 균형 로지스틱 회귀', 'features': 13, 'threshold_candidates': 33, 'selected_threshold': 0.175, 'holdout_used_for_training_or_selection': False}, 'goal': '후보 수 23개는 유지하면서 후보 내부 최저 확률 번호 1개와 후보 밖 최고 확률 번호 1개를 선택적으로 교체', 'baseline_23': {'rounds': 86, 'average_hits': 3.0232558139534884, 'three_plus_rate': 0.686046511627907, 'four_plus_rate': 0.32558139534883723, 'five_plus_rate': 0.09302325581395349, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'boundary_swap_23': {'rounds': 86, 'changed_rounds': 5, 'average_hits': 3.0232558139534884, 'three_plus_rate': 0.686046511627907, 'four_plus_rate': 0.32558139534883723, 'five_plus_rate': 0.09302325581395349, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'change': {'average_hits': 0.0, 'three_plus_pp': 0.0, 'four_plus_pp': 0.0, 'five_plus_pp': 0.0, 'six_all_pp': 0.0, 'six_all_cases': 0}, 'decision': {'status': '폐기', 'champion': 'v8.3 이중후보 23개 공격형 보조 챔피언 유지', 'reason': '보류검증에서 6개 전부 포함이 3회에서 3회로 감소했고 평균 포함도 3.023개에서 3.023개로 하락했습니다.', 'next_research': '후보 교체를 매 회차 실행하지 않고 회차 유형별 전문가 엔진이 합의할 때만 교체하는 다중 동의 방식'}, 'latest': {'baseline23': [15, 27, 38, 31, 3, 35, 10, 45, 20, 19, 44, 30, 17, 33, 29, 7, 8, 23, 37, 13, 34, 16, 6], 'remove_candidate': 6, 'add_candidate': 22, 'score_delta': 0.104333, 'threshold': 0.175, 'swap_applied': False, 'boundary_swap23': [15, 27, 38, 31, 3, 35, 10, 45, 20, 19, 44, 30, 17, 33, 29, 7, 8, 23, 37, 13, 34, 16, 6]}}


class TKMultiExpertConsensusLab:
    RESULT = {'version': '8.7.0-multi-expert-consensus', 'protocol': {'initial_training': '1~1000회', 'training': '1001~1088회', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'experts': ['최근형', '장기형', '패턴DNA형'], 'expert_feature_counts': {'최근형': 8, '장기형': 9, '패턴DNA형': 8}, 'selected_min_votes': 2, 'selected_threshold': 0.12, 'search_cases': 62, 'holdout_used_for_training_or_selection': False}, 'goal': '최근형·장기형·패턴DNA형 전문가가 같은 제거·추가 번호에 합의할 때만 후보23 경계번호 교체', 'baseline_23': {'rounds': 86, 'average_hits': 3.0232558139534884, 'three_plus_rate': 0.686046511627907, 'four_plus_rate': 0.32558139534883723, 'five_plus_rate': 0.09302325581395349, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'consensus_swap_23': {'rounds': 86, 'changed_rounds': 16, 'average_hits': 3.0348837209302326, 'three_plus_rate': 0.6744186046511628, 'four_plus_rate': 0.3372093023255814, 'five_plus_rate': 0.10465116279069768, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'change': {'average_hits': 0.0116, 'three_plus_pp': -1.16, 'four_plus_pp': 1.16, 'five_plus_pp': 1.16, 'six_all_pp': 0.0, 'six_all_cases': 0}, 'decision': {'status': '조건부 채택', 'champion': 'v8.7 다중전문가 합의형 보조엔진', 'reason': '6개 전부 포함을 유지하거나 늘리면서 5개 이상 및 평균 포함이 개선됐습니다.', 'next_research': '합의 대상 번호가 같을 때뿐 아니라 후보 경계군 3개와 후보 밖 경계군 3개의 집합 투표 방식 연구'}, 'latest': {'baseline23': [15, 27, 38, 31, 3, 35, 10, 45, 20, 19, 44, 30, 17, 33, 29, 7, 8, 23, 37, 13, 34, 16, 6], 'expert_proposals': {'최근형': {'remove': 6, 'add': 43, 'delta': 0.1226620627940147}, '장기형': {'remove': 34, 'add': 22, 'delta': 0.08430388595473431}, '패턴DNA형': {'remove': 6, 'add': 22, 'delta': 0.09878969328503617}}, 'consensus_action': None, 'consensus23': [15, 27, 38, 31, 3, 35, 10, 45, 20, 19, 44, 30, 17, 33, 29, 7, 8, 23, 37, 13, 34, 16, 6]}}


class TKBoundaryGroupVoteLab:
    RESULT = {'version': '8.8.0-boundary-group-vote', 'protocol': {'initial_training': '1~1000회', 'training': '1001~1088회', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'experts': ['최근형', '장기형', '패턴DNA형'], 'group_size': 3, 'selected_min_vote_score': 4, 'selected_threshold': 0.12, 'search_cases': 186, 'holdout_used_for_training_or_selection': False}, 'goal': '각 전문가의 후보 내부 하위3개와 후보 밖 상위3개를 집합 투표해 경계번호 1개를 선택 교체', 'baseline_23': {'rounds': 86, 'average_hits': 3.0232558139534884, 'three_plus_rate': 0.686046511627907, 'four_plus_rate': 0.32558139534883723, 'five_plus_rate': 0.09302325581395349, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'group_vote_23': {'rounds': 86, 'changed_rounds': 28, 'average_hits': 3.046511627906977, 'three_plus_rate': 0.6744186046511628, 'four_plus_rate': 0.3372093023255814, 'five_plus_rate': 0.10465116279069768, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'change': {'average_hits': 0.0233, 'four_plus_pp': 1.16, 'five_plus_pp': 1.16, 'six_all_pp': 0.0, 'six_all_cases': 0}, 'decision': {'status': '조건부 채택', 'champion': 'v8.8 경계군 집합투표 보조엔진', 'reason': '6개 전부 포함을 유지하거나 늘리면서 5개 이상 및 평균 포함이 개선됐습니다.', 'next_research': '6개 전부 포함 성공회차에 특화된 성공회차 전문가와 일반 안정형 전문가를 분리한 이중 게이트'}}


class TKSuccessDualGateLab:
    RESULT = {'version': '8.9.0-success-dual-gate', 'protocol': {'initial_training': '1~1000회', 'training': '1001~1088회', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'base_engine': 'v8.8 경계군 집합투표', 'gate_model': '표준화 + 균형 로지스틱 회귀', 'gate_features': 11, 'selected_gate_threshold': 0.35, 'success_round_weight': 3.0, 'holdout_used_for_training_or_selection': False}, 'goal': 'v8.8 교체신호 중 성공회차 특성과 유사한 회차에서만 교체를 허용하는 이중 게이트', 'baseline_23': {'rounds': 86, 'average_hits': 3.0232558139534884, 'three_plus_rate': 0.686046511627907, 'four_plus_rate': 0.32558139534883723, 'five_plus_rate': 0.09302325581395349, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'success_dual_gate_23': {'rounds': 86, 'changed_rounds': 10, 'average_hits': 3.058139534883721, 'three_plus_rate': 0.6976744186046512, 'four_plus_rate': 0.32558139534883723, 'five_plus_rate': 0.10465116279069768, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'change': {'average_hits': 0.0349, 'four_plus_pp': 0.0, 'five_plus_pp': 1.16, 'six_all_pp': 0.0, 'six_all_cases': 0}, 'decision': {'status': '조건부 채택', 'champion': 'v8.9 성공회차 이중게이트 보조엔진', 'reason': '6개 전부 포함을 유지하거나 늘리면서 5개 이상과 평균 포함이 함께 개선됐습니다.', 'next_research': '성공회차 가중치를 고정하지 않고 4·5·6개 적중 단계별로 다르게 주는 다단계 성공게이트'}}


class TKMultistageSuccessGateLab:
    RESULT = {'version': '9.0.0-multistage-success-gate', 'protocol': {'initial_training': '1~1000회', 'training': '1001~1088회', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'base_engine': 'v8.8 경계군 집합투표', 'gate_model': '표준화 + 균형 로지스틱 회귀', 'gate_features': 13, 'selected_stage_weights': {'four_plus': 1.2, 'five_plus': 2.0, 'six_all': 4.0}, 'selected_gate_threshold': 0.3, 'search_cases': 115, 'holdout_used_for_training_or_selection': False}, 'goal': '4개·5개·6개 적중 단계를 각각 다르게 가중한 다단계 성공게이트', 'baseline_23': {'rounds': 86, 'average_hits': 3.0232558139534884, 'three_plus_rate': 0.686046511627907, 'four_plus_rate': 0.32558139534883723, 'five_plus_rate': 0.09302325581395349, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'multistage_success_gate_23': {'rounds': 86, 'changed_rounds': 9, 'average_hits': 3.058139534883721, 'three_plus_rate': 0.6976744186046512, 'four_plus_rate': 0.32558139534883723, 'five_plus_rate': 0.10465116279069768, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'change': {'average_hits': 0.0349, 'four_plus_pp': 0.0, 'five_plus_pp': 1.16, 'six_all_pp': 0.0, 'six_all_cases': 0}, 'decision': {'status': '조건부 채택', 'champion': 'v9.0 다단계 성공게이트 보조엔진', 'reason': '6개 전부 포함을 유지하거나 늘리면서 5개 이상과 평균 포함이 함께 개선됐습니다.', 'next_research': '후보23을 유지한 채 조합 AI가 각 후보번호의 노출빈도를 적중단계별로 최적화하는 분배엔진'}}


class TKCombinationExposureLab:
    RESULT = {'version': '9.1.0-combination-exposure', 'protocol': {'inherited_training': '1~1088회 기존 후보엔진', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'candidate_size': 23, 'combinations_per_round': 100, 'candidate_configurations': 3, 'selected_alpha': 1.2, 'selected_exposure_weight': 0.7, 'selected_overlap_weight': 0.5, 'holdout_used_for_selection': False}, 'goal': '후보23에서 100조합의 번호 노출빈도와 조합 중복도를 최적화', 'baseline_top100': {'rounds': 86, 'best_hit_average': 1.9069767441860466, 'average_combo_hits': 0.7972093023255814, 'three_plus_rate': 0.23255813953488372, 'four_plus_rate': 0.05813953488372093, 'five_plus_rate': 0.0, 'six_all_rate': 0.0, 'six_all_cases': 0, 'max_hits': 4}, 'exposure_optimized_100': {'rounds': 86, 'best_hit_average': 1.9186046511627908, 'average_combo_hits': 0.7965116279069767, 'three_plus_rate': 0.2558139534883721, 'four_plus_rate': 0.046511627906976744, 'five_plus_rate': 0.0, 'six_all_rate': 0.0, 'six_all_cases': 0, 'max_hits': 4}, 'change': {'best_hit_average': 0.0116, 'four_plus_pp': -1.16, 'five_plus_pp': 0.0, 'six_all_cases': 0}, 'decision': {'status': '폐기', 'champion': 'v9.0 후보엔진 + 기존 조합배치', 'reason': '보류검증에서 기존 상위점수 100조합보다 종합적으로 우수하지 않아 채택하지 않았습니다.', 'next_research': '100조합을 안정형60·공격형40으로 분리한 이중 포트폴리오 조합엔진'}}


class TKDualPortfolioLab:
    RESULT = {'version': '9.2.0-dual-portfolio', 'protocol': {'inherited_training': '1~1088회 기존 후보엔진', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'candidate_size': 23, 'combinations_per_round': 100, 'tested_configurations': 4, 'selected_stable_count': 70, 'selected_aggressive_count': 30, 'selected_overlap_limit': 4, 'holdout_used_for_selection': False}, 'goal': '100조합을 안정형과 공격형으로 나눠 평균 적중과 고적중을 동시에 개선', 'baseline_top100': {'rounds': 86, 'best_hit_average': 1.9069767441860466, 'average_combo_hits': 0.7972093023255814, 'three_plus_rate': 0.23255813953488372, 'four_plus_rate': 0.05813953488372093, 'five_plus_rate': 0.0, 'six_all_rate': 0.0, 'six_all_cases': 0, 'max_hits': 4}, 'dual_portfolio_100': {'rounds': 86, 'best_hit_average': 2.0697674418604652, 'average_combo_hits': 0.7836046511627907, 'three_plus_rate': 0.3023255813953488, 'four_plus_rate': 0.06976744186046512, 'five_plus_rate': 0.0, 'six_all_rate': 0.0, 'six_all_cases': 0, 'max_hits': 4}, 'change': {'best_hit_average': 0.1628, 'average_combo_hits': -0.0136, 'four_plus_pp': 1.16, 'five_plus_pp': 0.0, 'six_all_cases': 0}, 'decision': {'status': '조건부 채택', 'champion': 'v9.2 이중 포트폴리오 조합엔진', 'reason': '6개 적중을 유지하거나 늘리면서 4개 이상과 최고조합 평균이 함께 개선됐습니다.', 'next_research': '안정형·공격형 조합 비중을 회차 유형별로 바꾸는 동적 포트폴리오 게이트'}}


class TKDynamicPortfolioLab:
    RESULT = {'version': '9.3.0-dynamic-portfolio', 'protocol': {'training': '1001~1088회', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'round_type_features': 16, 'portfolio_options': {'0': '70:30', '1': '60:40', '2': '50:50'}, 'tested_models': ['logistic_C0.3', 'logistic_C1.0', 'rf_depth4'], 'selected_model': 'rf_depth4', 'holdout_used_for_selection': False}, 'goal': '직전 회차와 최근 5·10회 특성으로 안정형·공격형 비중을 회차별 자동 선택', 'baseline_fixed_70_30': {'rounds': 86, 'best_hit_average': 2.0697674418604652, 'average_combo_hits': 0.7836046511627907, 'three_plus_rate': 0.3023255813953488, 'four_plus_rate': 0.06976744186046512, 'five_plus_rate': 0.0, 'six_all_rate': 0.0, 'six_all_cases': 0, 'max_hits': 4}, 'dynamic_portfolio': {'rounds': 86, 'best_hit_average': 2.058139534883721, 'average_combo_hits': 0.7873255813953488, 'three_plus_rate': 0.29069767441860467, 'four_plus_rate': 0.06976744186046512, 'five_plus_rate': 0.0, 'six_all_rate': 0.0, 'six_all_cases': 0, 'max_hits': 4, 'portfolio_counts': {'2': 21, '1': 33, '0': 32}}, 'change': {'best_hit_average': -0.0116, 'four_plus_pp': 0.0, 'five_plus_pp': 0.0, 'six_all_cases': 0}, 'decision': {'status': '폐기', 'champion': 'v9.2 고정 70:30 이중 포트폴리오', 'reason': '회차 유형 기반 동적 비중 선택이 고정 70:30 포트폴리오보다 보류검증에서 종합 우위를 만들지 못했습니다.', 'next_research': '회차 유형 분류 정확도를 먼저 검증하고 유형별 전용 후보엔진까지 분리하는 2단계 구조'}}


class TKFailureRecoveryLab:
    RESULT = {'version': '9.4.0-failure-recovery', 'protocol': {'training': '1001~1088회', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'features': 14, 'tested_settings': 45, 'selected_swap_count': 1, 'selected_threshold': 0.18, 'holdout_used_for_selection': False}, 'goal': '후보 밖 당첨번호와 후보 안 오선정 번호를 학습해 후보23을 최대 3개까지 선택 교체', 'baseline_candidate23': {'rounds': 86, 'average_hits': 3.0232558139534884, 'three_plus_rate': 0.686046511627907, 'four_plus_rate': 0.32558139534883723, 'five_plus_rate': 0.09302325581395349, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'failure_recovery_candidate23': {'rounds': 86, 'total_swaps': 4, 'average_hits': 3.0232558139534884, 'three_plus_rate': 0.686046511627907, 'four_plus_rate': 0.32558139534883723, 'five_plus_rate': 0.09302325581395349, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'change': {'average_hits': 0.0, 'four_plus_pp': 0.0, 'five_plus_pp': 0.0, 'six_all_cases': 0}, 'decision': {'status': '폐기', 'champion': 'v9.0 다단계 성공게이트 후보엔진', 'reason': '실패회차 복구 교체가 기존 후보23보다 보류검증에서 종합적인 개선을 만들지 못했습니다.', 'next_research': '후보 밖 번호를 한 번에 교체하지 않고 회차유형별 실패원인 전문가가 합의할 때만 복구하는 다중 실패전문가 구조'}}


class TKMultiFailureExpertsLab:
    RESULT = {'version': '9.5.0-multi-failure-experts', 'protocol': {'training': '1001~1088회', 'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'experts': ['최근실패형', '장기실패형', '패턴실패형'], 'tested_settings': 22, 'selected_min_votes': 3, 'selected_threshold': 0.12, 'holdout_used_for_selection': False}, 'goal': '최근·장기·패턴 실패전문가가 같은 제거·추가 후보에 합의할 때만 후보23 경계번호 교체', 'baseline_candidate23': {'rounds': 86, 'average_hits': 3.0232558139534884, 'three_plus_rate': 0.686046511627907, 'four_plus_rate': 0.32558139534883723, 'five_plus_rate': 0.09302325581395349, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'multi_failure_candidate23': {'rounds': 86, 'changed_rounds': 23, 'average_hits': 3.058139534883721, 'three_plus_rate': 0.686046511627907, 'four_plus_rate': 0.3372093023255814, 'five_plus_rate': 0.09302325581395349, 'six_all_rate': 0.03488372093023256, 'six_all_cases': 3, 'max_hits': 6}, 'change': {'average_hits': 0.0349, 'four_plus_pp': 1.16, 'five_plus_pp': 0.0, 'six_all_cases': 0}, 'decision': {'status': '조건부 채택', 'champion': 'v9.5 다중 실패전문가 후보엔진', 'reason': '다중 실패전문가 합의가 6개 포함을 유지하거나 늘리면서 평균 포함과 고적중 지표를 개선했습니다.', 'next_research': '후보23 전체를 교체하지 않고 후보 밖 번호의 복구 확률만 별도 산출해 보조후보 2개로 유지하는 안전 복구층'}}


class TKSafeRecoveryLayerLab:
    RESULT = {'version': '9.6.0-safe-recovery-layer', 'protocol': {'validation': '1089~1132회', 'holdout': '1133~1218회 (86회)', 'stable_combinations': 70, 'aggressive_combinations': 30, 'tested_aux_counts': [1, 2, 3], 'selected_aux_count': 2, 'base_candidate_size': 23, 'holdout_used_for_selection': False}, 'goal': '기존 후보23을 보존하고 후보 밖 보조번호를 공격형30조합에만 제한적으로 사용', 'baseline_70_30': {'rounds': 86, 'best_hit_average': 2.13953488372093, 'average_combo_hits': 0.781860465116279, 'three_plus_rate': 0.32558139534883723, 'four_plus_rate': 0.08139534883720931, 'five_plus_rate': 0.011627906976744186, 'six_all_rate': 0.0, 'six_all_cases': 0, 'max_hits': 5}, 'safe_recovery_70_30': {'rounds': 86, 'best_hit_average': 2.1627906976744184, 'average_combo_hits': 0.7888372093023258, 'three_plus_rate': 0.3372093023255814, 'four_plus_rate': 0.08139534883720931, 'five_plus_rate': 0.011627906976744186, 'six_all_rate': 0.0, 'six_all_cases': 0, 'max_hits': 5}, 'change': {'best_hit_average': 0.0233, 'four_plus_pp': 0.0, 'five_plus_pp': 0.0, 'six_all_cases': 0}, 'decision': {'status': '폐기', 'champion': 'v9.2 고정70:30 조합엔진 + v9.5 후보엔진', 'reason': '안전 복구층이 기존 70:30 조합엔진보다 보류검증에서 종합 우위를 만들지 못했습니다.', 'next_research': '보조후보를 동반출현 페어 단위로 투입하고 공격형30의 보조번호 노출 상한을 제한'}}


class TKRegimeCandidateLab:
    RESULT = {'version': '10.0.0-regime-candidate-engine', 'protocol': {'feature_start_round': 301, 'training': '301~1000회', 'validation': '1001~1100회', 'holdout': '1101~1218회 (118회)', 'candidate_size': 23, 'features': 9, 'regimes': 4, 'tested_regime_weights': [0.25, 0.5, 0.75, 1.0], 'selected_regime_weight': 0.5, 'holdout_used_for_selection': False}, 'goal': '회차유형을 먼저 분류하고 유형별 전용 번호모델과 글로벌 모델을 결합해 후보23 생성', 'baseline_global_candidate23': {'rounds': 118, 'average_hits': 3.0084745762711864, 'three_plus_rate': 0.6779661016949152, 'four_plus_rate': 0.3305084745762712, 'five_plus_rate': 0.09322033898305085, 'six_all_rate': 0.00847457627118644, 'six_all_cases': 1, 'max_hits': 6}, 'regime_candidate23': {'rounds': 118, 'average_hits': 2.9745762711864407, 'three_plus_rate': 0.6779661016949152, 'four_plus_rate': 0.3474576271186441, 'five_plus_rate': 0.07627118644067797, 'six_all_rate': 0.0, 'six_all_cases': 0, 'max_hits': 5}, 'change': {'average_hits': -0.0339, 'four_plus_pp': 1.69, 'five_plus_pp': -1.69, 'six_all_cases': -1}, 'decision': {'status': '폐기', 'champion': '기존 v9.5 후보엔진 유지', 'reason': '회차유형 전용 모델이 글로벌 후보모델보다 완전 보류검증에서 종합 우위를 만들지 못했습니다.', 'next_research': '회차유형 분류를 4종에서 6~8종으로 세분화하고 각 유형별 후보 수 20~25개를 별도 최적화'}}


class TKRegimeCandidateSizeLab:
    RESULT = {'version': '10.1.0-regime-candidate-size', 'protocol': {'training': '301~1000회', 'validation': '1001~1100회', 'holdout': '1101~1218회 (118회)', 'regimes': 4, 'candidate_sizes_tested': [20, 21, 22, 23, 24, 25], 'selected_sizes': {'0': 25, '1': 25, '2': 25, '3': 23}, 'holdout_used_for_selection': False}, 'goal': '회차유형별 후보 수를 20~25개에서 독립 선택해 후보 포함률과 효율을 동시에 개선', 'baseline_fixed23': {'rounds': 118, 'average_candidate_size': 23.0, 'average_hits': 2.9745762711864407, 'three_plus_rate': 0.6779661016949152, 'four_plus_rate': 0.3474576271186441, 'five_plus_rate': 0.07627118644067797, 'six_all_rate': 0.0, 'six_all_cases': 0, 'coverage_efficiency': 0.12932940309506263, 'max_hits': 5}, 'regime_dynamic_size': {'rounds': 118, 'average_candidate_size': 24.71186440677966, 'average_hits': 3.1779661016949152, 'three_plus_rate': 0.7457627118644068, 'four_plus_rate': 0.4322033898305085, 'five_plus_rate': 0.0847457627118644, 'six_all_rate': 0.00847457627118644, 'six_all_cases': 1, 'coverage_efficiency': 0.1286008230452675, 'max_hits': 6}, 'selection_detail': {'0': {'selected_size': 25, 'validation': {'rounds': 35, 'average_candidate_size': 25.0, 'average_hits': 3.342857142857143, 'three_plus_rate': 0.6857142857142857, 'four_plus_rate': 0.5428571428571428, 'five_plus_rate': 0.17142857142857143, 'six_all_rate': 0.02857142857142857, 'six_all_cases': 1, 'coverage_efficiency': 0.13371428571428573, 'max_hits': 6}}, '1': {'selected_size': 25, 'validation': {'rounds': 10, 'average_candidate_size': 25.0, 'average_hits': 3.2, 'three_plus_rate': 0.6, 'four_plus_rate': 0.5, 'five_plus_rate': 0.2, 'six_all_rate': 0.0, 'six_all_cases': 0, 'coverage_efficiency': 0.128, 'max_hits': 5}}, '2': {'selected_size': 25, 'validation': {'rounds': 40, 'average_candidate_size': 25.0, 'average_hits': 3.325, 'three_plus_rate': 0.75, 'four_plus_rate': 0.4, 'five_plus_rate': 0.2, 'six_all_rate': 0.05, 'six_all_cases': 2, 'coverage_efficiency': 0.133, 'max_hits': 6}}, '3': {'selected_size': 23, 'validation': {'rounds': 15, 'average_candidate_size': 23.0, 'average_hits': 3.8666666666666667, 'three_plus_rate': 0.8666666666666667, 'four_plus_rate': 0.6666666666666666, 'five_plus_rate': 0.3333333333333333, 'six_all_rate': 0.0, 'six_all_cases': 0, 'coverage_efficiency': 0.1681159420289855, 'max_hits': 5}}}, 'change': {'average_candidate_size': 1.7119, 'average_hits': 0.2034, 'coverage_efficiency': -0.000729, 'four_plus_pp': 8.47, 'five_plus_pp': 0.85, 'six_all_cases': 1}, 'decision': {'status': '조건부 채택', 'champion': 'v10.1 유형별 후보 수 엔진', 'reason': '6개·5개 이상 성과를 보존하면서 평균 포함 또는 후보 효율이 개선됐습니다.', 'next_research': '유형을 6종으로 재정의하고 후보 수 선택을 적중률이 아니라 후보 밖 누락 원인과 함께 학습하는 메타게이트'}}


class TKEnsembleExclusionAILab:
    RESULT = {'version': '10.2.0-ensemble-exclusion-ai', 'protocol': {'training': '301~1000회', 'validation': '1001~1100회', 'holdout': '1101~1218회 (118회)', 'candidate_sizes': {'0': 25, '1': 25, '2': 25, '3': 23}, 'features': 17, 'tested_settings': 20, 'selected_tree_weight': 0.65, 'selected_exclude_penalty': 0.2, 'holdout_used_for_selection': False}, 'goal': '글로벌·회차유형·비선형 전문가와 제외 AI를 결합해 후보 밖 누락을 줄임', 'baseline_v10_1': {'rounds': 118, 'average_candidate_size': 24.71186440677966, 'average_hits': 3.1610169491525424, 'average_missed': 2.8389830508474576, 'three_plus_rate': 0.7457627118644068, 'four_plus_rate': 0.3983050847457627, 'five_plus_rate': 0.07627118644067797, 'six_all_rate': 0.01694915254237288, 'six_all_cases': 2, 'max_hits': 6}, 'ensemble_exclusion_ai': {'rounds': 118, 'average_candidate_size': 24.71186440677966, 'average_hits': 3.1610169491525424, 'average_missed': 2.8389830508474576, 'three_plus_rate': 0.6779661016949152, 'four_plus_rate': 0.3728813559322034, 'five_plus_rate': 0.15254237288135594, 'six_all_rate': 0.025423728813559324, 'six_all_cases': 3, 'max_hits': 6}, 'change': {'average_hits': 0.0, 'average_missed': 0.0, 'four_plus_pp': -2.54, 'five_plus_pp': 7.63, 'six_all_cases': 1}, 'decision': {'status': '폐기', 'champion': '안정형 v10.1 + 공격형 v10.2 이중 운영', 'reason': '평균 포함은 유지되고 5개 이상 적중률과 6개 전부 포함이 증가했습니다. 다만 4개 이상 적중률이 하락해 공격형 보조엔진으로만 채택합니다.', 'next_research': '현재 25개 후보를 핵심23+보조2로 분리하고 보조번호가 실제 후보 밖 누락을 복구하는 회차에서만 활성화하는 메타게이트'}}


class TKStableAggressiveMetaGateLab:
    RESULT = {'version': '10.3.0-stable-aggressive-meta-gate', 'protocol': {'base_training': '301~900회', 'meta_training': '901~1000회', 'validation': '1001~1100회', 'holdout': '1101~1218회 (118회)', 'candidate_sizes': {'0': 25, '1': 25, '2': 25, '3': 23}, 'gate_features': 21, 'tested_gate_models': ['logistic', 'rf'], 'tested_thresholds': 19, 'selected_gate_model': 'rf', 'selected_threshold': 0.55, 'holdout_used_for_selection': False}, 'goal': 'v10.1 안정형과 v10.2 공격형 중 이번 회차에 유리한 엔진을 메타게이트가 선택', 'stable_v10_1': {'rounds': 118, 'average_hits': 3.059322033898305, 'three_plus_rate': 0.6864406779661016, 'four_plus_rate': 0.3644067796610169, 'five_plus_rate': 0.09322033898305085, 'six_all_rate': 0.00847457627118644, 'six_all_cases': 1, 'max_hits': 6, 'choices': {'stable': 118, 'aggressive': 0}}, 'aggressive_v10_2': {'rounds': 118, 'average_hits': 3.1016949152542375, 'three_plus_rate': 0.711864406779661, 'four_plus_rate': 0.3389830508474576, 'five_plus_rate': 0.11016949152542373, 'six_all_rate': 0.01694915254237288, 'six_all_cases': 2, 'max_hits': 6, 'choices': {'stable': 0, 'aggressive': 118}}, 'meta_gate_v10_3': {'rounds': 118, 'average_hits': 3.059322033898305, 'three_plus_rate': 0.6864406779661016, 'four_plus_rate': 0.3644067796610169, 'five_plus_rate': 0.09322033898305085, 'six_all_rate': 0.00847457627118644, 'six_all_cases': 1, 'max_hits': 6, 'choices': {'stable': 117, 'aggressive': 1}}, 'change_vs_stable': {'average_hits': 0.0, 'four_plus_pp': 0.0, 'five_plus_pp': 0.0, 'six_all_cases': 0}, 'change_vs_aggressive': {'average_hits': -0.0424, 'four_plus_pp': 2.54, 'five_plus_pp': -1.69, 'six_all_cases': -1}, 'decision': {'status': '폐기', 'champion': '안정형 v10.1 + 공격형 v10.2 이중 운영 유지', 'reason': '메타게이트가 안정형 v10.1과 공격형 v10.2의 장점을 동시에 넘어서지 못해 정식 채택하지 않았습니다.', 'next_research': '메타게이트의 이진 선택 대신 안정형23 + 공격형 보조2를 합치는 핵심·보조 후보층으로 고적중과 4개 이상을 동시에 방어'}}


class TKCore23AuxLayerLab:
    RESULT = {'version': '10.4.0-core23-aux2-layer', 'protocol': {'training': '301~1000회', 'validation': '1001~1100회', 'holdout': '1101~1218회 (118회)', 'core_size': 23, 'tested_aux_counts': [1, 2], 'tested_min_gaps': [-0.1, -0.05, 0.0, 0.03, 0.06], 'selected_aux_count': 2, 'selected_min_gap': -0.1, 'holdout_used_for_selection': False}, 'goal': '안정형 핵심23을 보존하고 공격형 보조후보 최대2개를 별도 층으로 추가', 'stable_core23': {'rounds': 118, 'average_core_size': 23.0, 'average_aux_size': 1.0, 'average_total_size': 24.0, 'core_average_hits': 2.9237288135593222, 'average_hits': 3.0508474576271185, 'average_aux_hits': 0.1271186440677966, 'aux_hit_rate': 0.1271186440677966, 'three_plus_rate': 0.6864406779661016, 'four_plus_rate': 0.3728813559322034, 'five_plus_rate': 0.06779661016949153, 'six_all_rate': 0.00847457627118644, 'six_all_cases': 1, 'max_hits': 6}, 'aggressive25_reference': {'rounds': 118, 'average_hits': 3.211864406779661, 'four_plus_rate': 0.3813559322033898, 'five_plus_rate': 0.15254237288135594, 'six_all_rate': 0.01694915254237288, 'six_all_cases': 2}, 'core23_aux_layer': {'rounds': 118, 'average_core_size': 23.0, 'average_aux_size': 2.0, 'average_total_size': 25.0, 'core_average_hits': 2.9237288135593222, 'average_hits': 3.1779661016949152, 'average_aux_hits': 0.2542372881355932, 'aux_hit_rate': 0.23728813559322035, 'three_plus_rate': 0.711864406779661, 'four_plus_rate': 0.3983050847457627, 'five_plus_rate': 0.11864406779661017, 'six_all_rate': 0.01694915254237288, 'six_all_cases': 2, 'max_hits': 6}, 'change_vs_core23': {'average_hits': 0.1271, 'four_plus_pp': 2.54, 'five_plus_pp': 5.08, 'six_all_cases': 1}, 'decision': {'status': '폐기', 'champion': '안정형 v10.1 + 공격형 v10.2 이중 운영 유지', 'reason': '핵심23+보조후보층이 안정형 핵심23과 공격형25의 장점을 동시에 넘지 못했습니다.', 'next_research': '보조후보 2개를 최종 조합의 20~35%에만 노출시키는 조합 노출상한 엔진'}}


class TKSurvivalCompetitionLab:
    RESULT = {'version': '10.5.0-survival-competition-exclusion-ai', 'protocol': {'training': '301~900회', 'tuning': '901~1100회', 'holdout': '1101~1218회 (118회)', 'experts': ['글로벌', '회차유형', 'ExtraTrees', 'RandomForest', '제외Tree', '제외Logit'], 'tested_settings': 225, 'selected_survival_weights': [0.25, 0.25, 0.25, 0.25], 'selected_exclude_weight': 0.25, 'selected_candidate_size': 25, 'selected_hard_exclude': 0, 'holdout_used_for_selection': False}, 'goal': '여러 전문가가 번호별 생존점수를 투표하고 제외 AI 합의로 약한 번호를 제거', 'baseline_attack25': {'rounds': 118, 'average_candidate_size': 25.0, 'average_hits': 3.1864406779661016, 'average_missed': 2.8135593220338984, 'three_plus_rate': 0.7457627118644068, 'four_plus_rate': 0.3898305084745763, 'five_plus_rate': 0.1271186440677966, 'six_all_rate': 0.00847457627118644, 'six_all_cases': 1, 'max_hits': 6}, 'survival_competition': {'rounds': 118, 'average_candidate_size': 25.0, 'average_hits': 3.1016949152542375, 'average_missed': 2.8983050847457625, 'three_plus_rate': 0.7372881355932204, 'four_plus_rate': 0.3389830508474576, 'five_plus_rate': 0.1016949152542373, 'six_all_rate': 0.00847457627118644, 'six_all_cases': 1, 'max_hits': 6}, 'change': {'average_candidate_size': 0.0, 'average_hits': -0.0847, 'four_plus_pp': -5.08, 'five_plus_pp': -2.54, 'six_all_cases': 0}, 'decision': {'status': '폐기', 'champion': '안정형 v10.1 + 공격형 v10.2 유지', 'reason': '후보 생존경쟁·제외 AI가 기존 공격형 기준보다 완전 보류검증에서 핵심 지표를 개선하지 못했습니다.', 'next_research': '생존점수 상위 핵심번호와 경계번호를 분리하고 경계번호만 페어·트리플 동반출현 점수로 재심사'}}


class TKBoundaryCompanionReviewLab:
    RESULT = {'version': '10.6.0-boundary-companion-review', 'protocol': {'training': '1~900회 누적통계', 'tuning': '901~1100회', 'holdout': '1101~1218회 (118회)', 'tested_settings': 72, 'selected_boundary_width': 4, 'selected_pair_weight': 0.4, 'selected_triple_weight': 0.5, 'selected_candidate_size': 25, 'holdout_used_for_selection': False}, 'goal': '상위 핵심번호는 고정하고 경계번호만 페어·트리플 동반출현 점수로 재심사', 'baseline_attack25': {'rounds': 118, 'average_candidate_size': 25.0, 'average_hits': 3.26271186440678, 'three_plus_rate': 0.7372881355932204, 'four_plus_rate': 0.4406779661016949, 'five_plus_rate': 0.11864406779661017, 'six_all_rate': 0.05084745762711865, 'six_all_cases': 6, 'max_hits': 6, 'changed_rounds': 0, 'promoted_actual_hits': 0, 'demoted_actual_hits': 0, 'net_boundary_hits': 0}, 'boundary_review': {'rounds': 118, 'average_candidate_size': 25.0, 'average_hits': 3.2288135593220337, 'three_plus_rate': 0.7033898305084746, 'four_plus_rate': 0.423728813559322, 'five_plus_rate': 0.1440677966101695, 'six_all_rate': 0.0423728813559322, 'six_all_cases': 5, 'max_hits': 6, 'changed_rounds': 114, 'promoted_actual_hits': 17, 'demoted_actual_hits': 21, 'net_boundary_hits': -4}, 'change': {'candidate_size': 0.0, 'average_hits': -0.0339, 'four_plus_pp': -1.69, 'five_plus_pp': 2.54, 'six_all_cases': -1, 'net_boundary_hits': -4}, 'decision': {'status': '폐기', 'champion': '안정형 v10.1 + 공격형 v10.2 유지', 'reason': '경계번호 동반수·트리플 재심사가 기존 공격형25보다 완전 보류검증에서 핵심 성과를 개선하지 못했습니다.', 'next_research': '페어·트리플 전문가를 분리해 두 전문가가 모두 동의할 때만 경계번호를 승격하는 최소합의 엔진'}}


class TKDDayFinalLab:
    RESULT = {'version': '10.7.0-dday-final-combination', 'protocol': {'tuning': '1001~1100회', 'holdout': '1101~1218회 (118회)', 'presets_tested': ['균형형', '단기강화', '장기안정', '동반강화'], 'selected_preset': '장기안정', 'holdout_used_for_selection': False}, 'holdout': {'rounds': 118, 'candidate15_avg': 2.0084745762711864, 'candidate15_4plus': 0.11016949152542373, 'candidate15_5plus': 0.03389830508474576, 'candidate15_6all': 0, 'top5_best_avg': 1.271186440677966, 'top5_3plus': 0.07627118644067797, 'top5_4plus': 0.00847457627118644, 'top5_5plus': 0.0, 'top5_6all': 0}, 'next_round': 1219, 'top15': [27, 38, 15, 3, 19, 7, 24, 31, 37, 30, 35, 10, 16, 33, 9], 'top10': [27, 38, 15, 3, 19, 7, 24, 31, 37, 30], 'top7': [27, 38, 15, 3, 19, 7, 24], 'final5': [[27, 38, 15, 3, 24, 33], [27, 38, 15, 3, 19, 30], [27, 38, 15, 19, 31, 10], [27, 38, 15, 3, 37, 16], [27, 38, 15, 19, 24, 16]], 'disclaimer': '과거 데이터 기반 점수이며 미래 당첨을 보장하지 않습니다.'}


class TKFailureMemoryDynamicReliabilityLab:
    RESULT = {'version': '10.9.0-failure-memory-dynamic-reliability', 'protocol': {'expert_history_start': '901회', 'tuning': '1001~1100회', 'holdout': '1101~1218회 (118회)', 'experts': ['최근형', '중기형', '장기형', '미출현형', '구조형'], 'reliability_window': 60, 'failure_memory_window': 100, 'tested_gate_settings': 80, 'selected_margin23': 0.01, 'selected_margin24': 0.018, 'selected_disagreement_max': 0.16, 'holdout_used_for_selection': False}, 'goal': '엔진별 최근 신뢰도와 실패 회차 기억을 이용해 회차별 후보 수를 23~25개로 자동 결정', 'baseline_candidate25': {'rounds': 118, 'average_candidate_size': 25.0, 'average_hits': 3.3559322033898304, 'three_plus_rate': 0.7457627118644068, 'four_plus_rate': 0.4745762711864407, 'five_plus_rate': 0.211864406779661, 'six_all_rate': 0.01694915254237288, 'six_all_cases': 2, 'efficiency': 0.1342372881355932, 'size_choices': {23: 0, 24: 0, 25: 118}}, 'dynamic_failure_memory': {'rounds': 118, 'average_candidate_size': 24.915254237288135, 'average_hits': 3.3389830508474576, 'three_plus_rate': 0.7457627118644068, 'four_plus_rate': 0.4745762711864407, 'five_plus_rate': 0.2033898305084746, 'six_all_rate': 0.01694915254237288, 'six_all_cases': 2, 'efficiency': 0.1340136054421769, 'size_choices': {23: 0, 24: 10, 25: 108}}, 'change': {'average_candidate_size': -0.0847, 'average_hits': -0.0169, 'four_plus_pp': 0.0, 'five_plus_pp': -0.85, 'six_all_cases': 0, 'efficiency': -0.000224}, 'decision': {'status': '폐기', 'champion': '기존 후보25 유지', 'reason': '동적 신뢰도·실패기억 엔진이 후보25 대비 고적중 유지와 후보 축소를 동시에 충족하지 못했습니다.', 'next_research': '동적 후보수와 별도로 실제 누락 위험이 큰 회차에만 보조후보를 활성화하는 선택적 복구 게이트'}}


class TKLatest1232ProspectiveLab:
    RESULT = {'version': '11.0.0-latest-1232-prospective-validation', 'data': {'source_file': '로또 회차별 당첨번호_20260717113954.xlsx', 'rounds': 1232, 'latest_round': 1232, 'capture_cross_check': {1232: True, 1231: True, 1230: True, 1229: True, 1228: True, 1227: True, 1226: True, 1225: True, 1224: True, 1223: True, 1222: True, 1221: True, 1220: True, 1219: True}}, 'protocol': {'historical_design_end': '1218회', 'new_future_validation': '1219~1232회 (14회)', 'future_used_for_setting_selection': False, 'candidate_sizes_checked': [18, 19, 20, 21, 22, 23, 24, 25], 'v10_9_thresholds_fixed': [0.01, 0.018, 0.16]}, 'candidate25_future': {'rounds': 14, 'candidate_size': 25, 'average_hits': 2.7142857142857144, 'three_plus_rate': 0.5714285714285714, 'four_plus_rate': 0.35714285714285715, 'five_plus_rate': 0.0, 'six_all_cases': 0, 'efficiency': 0.10857142857142858}, 'dynamic_v10_9_future': {'rounds': 14, 'average_candidate_size': 24.928571428571427, 'average_hits': 3.0714285714285716, 'three_plus_rate': 0.6428571428571429, 'four_plus_rate': 0.2857142857142857, 'five_plus_rate': 0.21428571428571427, 'six_all_cases': 0, 'efficiency': 0.12320916905444128, 'size_choices': {'23': 0, '24': 1, '25': 13}}, 'all_candidate_sizes_future': {'18': {'rounds': 14, 'candidate_size': 18, 'average_hits': 1.9285714285714286, 'three_plus_rate': 0.2857142857142857, 'four_plus_rate': 0.14285714285714285, 'five_plus_rate': 0.0, 'six_all_cases': 0, 'efficiency': 0.10714285714285715}, '19': {'rounds': 14, 'candidate_size': 19, 'average_hits': 2.2142857142857144, 'three_plus_rate': 0.35714285714285715, 'four_plus_rate': 0.2857142857142857, 'five_plus_rate': 0.0, 'six_all_cases': 0, 'efficiency': 0.11654135338345865}, '20': {'rounds': 14, 'candidate_size': 20, 'average_hits': 2.2857142857142856, 'three_plus_rate': 0.35714285714285715, 'four_plus_rate': 0.2857142857142857, 'five_plus_rate': 0.0, 'six_all_cases': 0, 'efficiency': 0.11428571428571428}, '21': {'rounds': 14, 'candidate_size': 21, 'average_hits': 2.357142857142857, 'three_plus_rate': 0.42857142857142855, 'four_plus_rate': 0.2857142857142857, 'five_plus_rate': 0.0, 'six_all_cases': 0, 'efficiency': 0.11224489795918367}, '22': {'rounds': 14, 'candidate_size': 22, 'average_hits': 2.4285714285714284, 'three_plus_rate': 0.42857142857142855, 'four_plus_rate': 0.2857142857142857, 'five_plus_rate': 0.0, 'six_all_cases': 0, 'efficiency': 0.11038961038961038}, '23': {'rounds': 14, 'candidate_size': 23, 'average_hits': 2.5714285714285716, 'three_plus_rate': 0.5714285714285714, 'four_plus_rate': 0.2857142857142857, 'five_plus_rate': 0.0, 'six_all_cases': 0, 'efficiency': 0.11180124223602485}, '24': {'rounds': 14, 'candidate_size': 24, 'average_hits': 2.5714285714285716, 'three_plus_rate': 0.5714285714285714, 'four_plus_rate': 0.2857142857142857, 'five_plus_rate': 0.0, 'six_all_cases': 0, 'efficiency': 0.10714285714285715}, '25': {'rounds': 14, 'candidate_size': 25, 'average_hits': 2.7142857142857144, 'three_plus_rate': 0.5714285714285714, 'four_plus_rate': 0.35714285714285715, 'five_plus_rate': 0.0, 'six_all_cases': 0, 'efficiency': 0.10857142857142858}}, 'best_observed_size': 25, 'round_details_candidate25': [{'회차': 1219, '적중수': 4, '당첨번호': '1 2 15 28 39 45'}, {'회차': 1220, '적중수': 1, '당첨번호': '2 22 25 28 34 43'}, {'회차': 1221, '적중수': 4, '당첨번호': '6 13 18 28 30 36'}, {'회차': 1222, '적중수': 0, '당첨번호': '4 11 17 22 32 41'}, {'회차': 1223, '적중수': 3, '당첨번호': '16 18 20 32 33 39'}, {'회차': 1224, '적중수': 3, '당첨번호': '9 18 21 27 44 45'}, {'회차': 1225, '적중수': 3, '당첨번호': '8 9 19 25 41 42'}, {'회차': 1226, '적중수': 2, '당첨번호': '4 6 13 17 26 28'}, {'회차': 1227, '적중수': 2, '당첨번호': '1 14 16 34 41 44'}, {'회차': 1228, '적중수': 4, '당첨번호': '24 29 30 31 35 44'}, {'회차': 1229, '적중수': 2, '당첨번호': '12 13 29 34 37 42'}, {'회차': 1230, '적중수': 4, '당첨번호': '3 8 9 22 28 42'}, {'회차': 1231, '적중수': 4, '당첨번호': '4 13 14 18 31 38'}, {'회차': 1232, '적중수': 2, '당첨번호': '12 15 19 22 24 36'}], 'round_details_dynamic': [{'회차': 1219, '후보수': 25, '적중수': 4}, {'회차': 1220, '후보수': 25, '적중수': 2}, {'회차': 1221, '후보수': 25, '적중수': 5}, {'회차': 1222, '후보수': 24, '적중수': 2}, {'회차': 1223, '후보수': 25, '적중수': 2}, {'회차': 1224, '후보수': 25, '적중수': 5}, {'회차': 1225, '후보수': 25, '적중수': 3}, {'회차': 1226, '후보수': 25, '적중수': 2}, {'회차': 1227, '후보수': 25, '적중수': 3}, {'회차': 1228, '후보수': 25, '적중수': 5}, {'회차': 1229, '후보수': 25, '적중수': 3}, {'회차': 1230, '후보수': 25, '적중수': 3}, {'회차': 1231, '후보수': 25, '적중수': 1}, {'회차': 1232, '후보수': 25, '적중수': 3}], 'recommendation_1233': {'candidate25': [13, 27, 16, 28, 15, 31, 9, 18, 3, 24, 38, 44, 19, 42, 35, 12, 29, 22, 30, 8, 36, 6, 45, 37, 41], 'candidate15_reference': [13, 27, 16, 28, 15, 31, 9, 18, 3, 24, 38, 44, 19, 42, 35], 'note': '과거 데이터 기반 연구 후보이며 당첨을 보장하지 않음'}, 'decision': {'status': '관찰 유지', 'champion': '기존 후보25 유지', 'reason': '1219~1232회는 기존 연구 이후의 실제 미래 14회지만 표본이 작아 엔진을 교체하지 않고 후보25를 유지합니다.', 'next_research': '미래 회차가 30~50회 누적될 때 후보수와 동적게이트 재판정'}}


class TKRuleAblationCleanupLab:
    RESULT = {'version': '11.1.0-rule-ablation-cleanup', 'protocol': {'tuning': '1001~1100회', 'holdout': '1101~1218회 (118회)', 'prospective': '1219~1232회 (14회)', 'tested_variants': ['전체규칙', '최근제거', '중기제거', '장기제거', '간격제거', '이월인접제거', '순위안정제거', '핵심3종'], 'selected_on_tuning': '최근제거', 'holdout_used_for_selection': False, 'prospective_used_for_selection': False}, 'goal': '규칙을 하나씩 제거해 실제 기여도가 없는 요소를 정리', 'baseline_full': {'holdout': {'rounds': 118, 'candidate_size': 25, 'average_hits': 3.364406779661017, 'three_plus_rate': 0.7288135593220338, 'four_plus_rate': 0.4661016949152542, 'five_plus_rate': 0.2033898305084746, 'six_all_cases': 2, 'efficiency': 0.13457627118644067}, 'prospective': {'rounds': 14, 'candidate_size': 25, 'average_hits': 3.0, 'three_plus_rate': 0.6428571428571429, 'four_plus_rate': 0.5, 'five_plus_rate': 0.07142857142857142, 'six_all_cases': 0, 'efficiency': 0.12}}, 'selected_variant': {'name': '최근제거', 'weights': {'중기': 0.3125, '장기': 0.25, '간격': 0.125, '이월인접': 0.0625, '순위안정': 0.25}, 'holdout': {'rounds': 118, 'candidate_size': 25, 'average_hits': 3.457627118644068, 'three_plus_rate': 0.7627118644067796, 'four_plus_rate': 0.4745762711864407, 'five_plus_rate': 0.22033898305084745, 'six_all_cases': 5, 'efficiency': 0.13830508474576272}, 'prospective': {'rounds': 14, 'candidate_size': 25, 'average_hits': 3.2142857142857144, 'three_plus_rate': 0.6428571428571429, 'four_plus_rate': 0.5714285714285714, 'five_plus_rate': 0.21428571428571427, 'six_all_cases': 0, 'efficiency': 0.1285714285714286}}, 'all_variants': {'전체규칙': {'tuning': {'rounds': 100, 'candidate_size': 25, 'average_hits': 3.26, 'three_plus_rate': 0.76, 'four_plus_rate': 0.42, 'five_plus_rate': 0.14, 'six_all_cases': 1, 'efficiency': 0.1304}, 'holdout': {'rounds': 118, 'candidate_size': 25, 'average_hits': 3.364406779661017, 'three_plus_rate': 0.7288135593220338, 'four_plus_rate': 0.4661016949152542, 'five_plus_rate': 0.2033898305084746, 'six_all_cases': 2, 'efficiency': 0.13457627118644067}, 'prospective': {'rounds': 14, 'candidate_size': 25, 'average_hits': 3.0, 'three_plus_rate': 0.6428571428571429, 'four_plus_rate': 0.5, 'five_plus_rate': 0.07142857142857142, 'six_all_cases': 0, 'efficiency': 0.12}}, '최근제거': {'tuning': {'rounds': 100, 'candidate_size': 25, 'average_hits': 3.28, 'three_plus_rate': 0.73, 'four_plus_rate': 0.47, 'five_plus_rate': 0.13, 'six_all_cases': 3, 'efficiency': 0.13119999999999998}, 'holdout': {'rounds': 118, 'candidate_size': 25, 'average_hits': 3.457627118644068, 'three_plus_rate': 0.7627118644067796, 'four_plus_rate': 0.4745762711864407, 'five_plus_rate': 0.22033898305084745, 'six_all_cases': 5, 'efficiency': 0.13830508474576272}, 'prospective': {'rounds': 14, 'candidate_size': 25, 'average_hits': 3.2142857142857144, 'three_plus_rate': 0.6428571428571429, 'four_plus_rate': 0.5714285714285714, 'five_plus_rate': 0.21428571428571427, 'six_all_cases': 0, 'efficiency': 0.1285714285714286}}, '중기제거': {'tuning': {'rounds': 100, 'candidate_size': 25, 'average_hits': 3.31, 'three_plus_rate': 0.76, 'four_plus_rate': 0.45, 'five_plus_rate': 0.17, 'six_all_cases': 0, 'efficiency': 0.1324}, 'holdout': {'rounds': 118, 'candidate_size': 25, 'average_hits': 3.406779661016949, 'three_plus_rate': 0.7288135593220338, 'four_plus_rate': 0.4830508474576271, 'five_plus_rate': 0.1864406779661017, 'six_all_cases': 6, 'efficiency': 0.13627118644067795}, 'prospective': {'rounds': 14, 'candidate_size': 25, 'average_hits': 2.7142857142857144, 'three_plus_rate': 0.5714285714285714, 'four_plus_rate': 0.2857142857142857, 'five_plus_rate': 0.14285714285714285, 'six_all_cases': 0, 'efficiency': 0.10857142857142858}}, '장기제거': {'tuning': {'rounds': 100, 'candidate_size': 25, 'average_hits': 3.27, 'three_plus_rate': 0.77, 'four_plus_rate': 0.43, 'five_plus_rate': 0.14, 'six_all_cases': 1, 'efficiency': 0.1308}, 'holdout': {'rounds': 118, 'candidate_size': 25, 'average_hits': 3.4152542372881354, 'three_plus_rate': 0.7288135593220338, 'four_plus_rate': 0.4745762711864407, 'five_plus_rate': 0.2288135593220339, 'six_all_cases': 4, 'efficiency': 0.1366101694915254}, 'prospective': {'rounds': 14, 'candidate_size': 25, 'average_hits': 2.7857142857142856, 'three_plus_rate': 0.6428571428571429, 'four_plus_rate': 0.2857142857142857, 'five_plus_rate': 0.07142857142857142, 'six_all_cases': 0, 'efficiency': 0.11142857142857142}}, '간격제거': {'tuning': {'rounds': 100, 'candidate_size': 25, 'average_hits': 3.3, 'three_plus_rate': 0.8, 'four_plus_rate': 0.41, 'five_plus_rate': 0.15, 'six_all_cases': 1, 'efficiency': 0.132}, 'holdout': {'rounds': 118, 'candidate_size': 25, 'average_hits': 3.3728813559322033, 'three_plus_rate': 0.7288135593220338, 'four_plus_rate': 0.4576271186440678, 'five_plus_rate': 0.2033898305084746, 'six_all_cases': 4, 'efficiency': 0.13491525423728812}, 'prospective': {'rounds': 14, 'candidate_size': 25, 'average_hits': 3.0, 'three_plus_rate': 0.6428571428571429, 'four_plus_rate': 0.42857142857142855, 'five_plus_rate': 0.14285714285714285, 'six_all_cases': 0, 'efficiency': 0.12}}, '이월인접제거': {'tuning': {'rounds': 100, 'candidate_size': 25, 'average_hits': 3.29, 'three_plus_rate': 0.79, 'four_plus_rate': 0.41, 'five_plus_rate': 0.16, 'six_all_cases': 1, 'efficiency': 0.1316}, 'holdout': {'rounds': 118, 'candidate_size': 25, 'average_hits': 3.3983050847457625, 'three_plus_rate': 0.7457627118644068, 'four_plus_rate': 0.4661016949152542, 'five_plus_rate': 0.211864406779661, 'six_all_cases': 3, 'efficiency': 0.1359322033898305}, 'prospective': {'rounds': 14, 'candidate_size': 25, 'average_hits': 2.9285714285714284, 'three_plus_rate': 0.6428571428571429, 'four_plus_rate': 0.42857142857142855, 'five_plus_rate': 0.07142857142857142, 'six_all_cases': 0, 'efficiency': 0.11714285714285713}}, '순위안정제거': {'tuning': {'rounds': 100, 'candidate_size': 25, 'average_hits': 3.27, 'three_plus_rate': 0.78, 'four_plus_rate': 0.42, 'five_plus_rate': 0.13, 'six_all_cases': 1, 'efficiency': 0.1308}, 'holdout': {'rounds': 118, 'candidate_size': 25, 'average_hits': 3.3813559322033897, 'three_plus_rate': 0.7372881355932204, 'four_plus_rate': 0.4745762711864407, 'five_plus_rate': 0.211864406779661, 'six_all_cases': 2, 'efficiency': 0.1352542372881356}, 'prospective': {'rounds': 14, 'candidate_size': 25, 'average_hits': 3.0, 'three_plus_rate': 0.6428571428571429, 'four_plus_rate': 0.5, 'five_plus_rate': 0.07142857142857142, 'six_all_cases': 0, 'efficiency': 0.12}}, '핵심3종': {'tuning': {'rounds': 100, 'candidate_size': 25, 'average_hits': 3.29, 'three_plus_rate': 0.77, 'four_plus_rate': 0.47, 'five_plus_rate': 0.1, 'six_all_cases': 1, 'efficiency': 0.1316}, 'holdout': {'rounds': 118, 'candidate_size': 25, 'average_hits': 3.3983050847457625, 'three_plus_rate': 0.7711864406779662, 'four_plus_rate': 0.4406779661016949, 'five_plus_rate': 0.22033898305084745, 'six_all_cases': 4, 'efficiency': 0.1359322033898305}, 'prospective': {'rounds': 14, 'candidate_size': 25, 'average_hits': 2.7142857142857144, 'three_plus_rate': 0.6428571428571429, 'four_plus_rate': 0.2857142857142857, 'five_plus_rate': 0.07142857142857142, 'six_all_cases': 0, 'efficiency': 0.10857142857142858}}}, 'decision': {'status': '조건부 채택', 'champion': '최근제거', 'reason': '최근제거이 조정검증에서 선택됐고 보류검증과 실제 미래 14회에서도 전체규칙을 밀어내지 않았습니다.', 'next_research': '1233회 이후 실제 미래 회차를 계속 누적하며 규칙별 기여도를 재평가'}, 'recommendation_1233': {'candidate25': [13, 16, 3, 27, 12, 6, 38, 35, 15, 28, 7, 30, 33, 24, 45, 31, 9, 37, 29, 21, 19, 18, 36, 26, 1], 'top15_reference': [13, 16, 3, 27, 12, 6, 38, 35, 15, 28, 7, 30, 33, 24, 45], 'note': '과거 데이터 기반 연구 후보이며 당첨을 보장하지 않음'}}

class Recommender:
    """입력빈도·동반수·트리플·최근패턴을 자동 종합해 순위를 계산합니다."""

    CATEGORY_NAMES = {
        "추천조합": "composite",
        "나온횟수": "input",
        "동반수": "pair",
        "트리플": "triple",
        "최근패턴": "recent",
        "통합데이터추천": "mixed",
        "성과최적추천": "performance",
        "AI진화추천": "evolution",
        "AI연구소추천": "evolution_lab",
        "AI후보번호연구": "candidate_lab",
        "AI이중후보추천": "dual_candidate",
        "특이패턴추천": "pattern",
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

    def triple_details(self, combo: tuple[int, ...], top_n: int = 3):
        details = [
            (triple, self.a.triple_counts[triple])
            for triple in combinations(combo, 3)
        ]
        details.sort(key=lambda item: (-item[1], item[0]))
        return details[:top_n]

    @staticmethod
    def confidence_score(score: float, metrics: dict[str, float]) -> float:
        """0~100 추천 신뢰도. 통계점수와 조합 균형을 함께 반영합니다."""
        base = min(100.0, max(0.0, score))
        stability = (
            metrics.get("pair", 0.0) * 0.25
            + metrics.get("triple", 0.0) * 0.20
            + metrics.get("recent", 0.0) * 0.20
            + metrics.get("structure", 0.0) * 0.20
            + metrics.get("input", 0.0) * 0.15
        )
        return round(min(100.0, base * 0.65 + stability * 0.35), 1)

    @staticmethod
    def confidence_grade(confidence: float) -> str:
        if confidence >= 95:
            return "S"
        if confidence >= 90:
            return "A"
        if confidence >= 80:
            return "B"
        if confidence >= 70:
            return "C"
        return "D"

    @staticmethod
    def select_diverse(
        candidates: list[tuple[float, tuple[int, ...], dict[str, float]]],
        count: int,
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        """v40 Elite Survival: 점수·군집다양성·번호노출 균형으로 최종 조합을 생존시킵니다."""
        if not candidates or count <= 0:
            return []

        # 상위 후보군을 충분히 남긴 뒤 조합공간을 탐색합니다.
        working = sorted(candidates, key=lambda row: (-row[0], row[1]))[:max(count * 35, 500)]
        selected: list[tuple[float, tuple[int, ...], dict[str, float]]] = []
        selected_sets: list[set[int]] = []
        number_usage: Counter[int] = Counter()
        zone_usage: Counter[tuple[int, ...]] = Counter()

        def signature(combo: tuple[int, ...]) -> tuple[int, ...]:
            return tuple(sum(lo <= n <= hi for n in combo) for lo, hi in (
                (1, 10), (11, 20), (21, 30), (31, 40), (41, 45)
            ))

        while working and len(selected) < count:
            best_index = -1
            best_value = float('-inf')
            for index, (base_score, combo, metrics) in enumerate(working):
                combo_set = set(combo)
                max_overlap = max((len(combo_set & old) for old in selected_sets), default=0)
                overlap_penalty = {0: 0.0, 1: 0.0, 2: 0.5, 3: 2.0, 4: 8.0, 5: 24.0, 6: 100.0}[max_overlap]
                exposure_penalty = sum(number_usage[n] ** 1.25 for n in combo) * 0.42
                sig = signature(combo)
                cluster_penalty = zone_usage[sig] * 2.8
                underused_bonus = sum(1.0 / (1.0 + number_usage[n]) for n in combo) * 1.8
                elite_value = base_score + underused_bonus - overlap_penalty - exposure_penalty - cluster_penalty
                if elite_value > best_value:
                    best_value = elite_value
                    best_index = index

            base_score, combo, metrics = working.pop(best_index)
            if combo in {row[1] for row in selected}:
                continue
            enriched = dict(metrics)
            enriched['combination_engine'] = 'v40 Elite Survival'
            enriched['elite_survival_score'] = round(best_value, 4)
            selected.append((base_score, combo, enriched))
            selected_sets.append(set(combo))
            number_usage.update(combo)
            zone_usage[signature(combo)] += 1

        return selected

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

    STRATEGY_WEIGHTS = {
        "균형형": {
            "input": 0.25, "pair": 0.20, "triple": 0.15,
            "recent": 0.15, "structure": 0.25,
        },
        "출현형": {
            "input": 0.50, "pair": 0.15, "triple": 0.10,
            "recent": 0.15, "structure": 0.10,
        },
        "동반수형": {
            "input": 0.15, "pair": 0.50, "triple": 0.15,
            "recent": 0.10, "structure": 0.10,
        },
        "트리플형": {
            "input": 0.10, "pair": 0.20, "triple": 0.50,
            "recent": 0.10, "structure": 0.10,
        },
        "최근형": {
            "input": 0.15, "pair": 0.15, "triple": 0.10,
            "recent": 0.50, "structure": 0.10,
        },
        "AI Ultimate": {
            "input": 0.20, "pair": 0.20, "triple": 0.15,
            "recent": 0.20, "structure": 0.25,
        },
    }

    def strategy_score(
        self,
        metrics: dict[str, float],
        strategy: str,
    ) -> float:
        weights = self.STRATEGY_WEIGHTS.get(
            strategy, self.STRATEGY_WEIGHTS["균형형"]
        )
        return sum(metrics.get(key, 0.0) * weight for key, weight in weights.items())

    MIXED_PRESETS = {
        "입력중심형": (0.50, 0.20, 0.20, 0.10),
        "최근중심형": (0.25, 0.40, 0.25, 0.10),
        "균형형": (0.30, 0.25, 0.25, 0.20),
        "장기형": (0.20, 0.15, 0.25, 0.40),
    }

    @staticmethod
    def _window_analyzer(draws: list[Draw], size: int) -> LottoAnalyzer:
        analyzer = LottoAnalyzer()
        analyzer.draws = list(draws[-min(size, len(draws)):])
        analyzer._analyze()
        return analyzer

    @staticmethod
    def _historical_combo_score(
        combo: tuple[int, ...],
        analyzer: LottoAnalyzer,
    ) -> float:
        max_number = max(analyzer.number_counts.values(), default=1)
        max_pair = max(analyzer.pair_counts.values(), default=1)
        max_triple = max(analyzer.triple_counts.values(), default=1)

        number_score = (
            sum(analyzer.number_counts[n] for n in combo)
            / max(1, max_number * 6)
            * 100
        )
        pair_values = sorted(
            (analyzer.pair_counts[p] for p in combinations(combo, 2)),
            reverse=True,
        )
        pair_score = sum(pair_values[:5]) / max(1, max_pair * 5) * 100

        triple_values = sorted(
            (analyzer.triple_counts[t] for t in combinations(combo, 3)),
            reverse=True,
        )
        triple_score = sum(triple_values[:5]) / max(1, max_triple * 5) * 100

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
        structure = max(0.0, min(100.0, structure))

        return (
            number_score * 0.35
            + pair_score * 0.30
            + triple_score * 0.20
            + structure * 0.15
        )

    def _auto_mixed_preset(
        self,
        source_weights: Counter[int],
    ) -> str:
        """최근 20회를 검증구간으로 사용해 네 혼합비율 중 하나를 빠르게 선택합니다."""
        if len(self.a.draws) < 120:
            return "균형형"

        history = self.a.draws[:-20]
        targets = self.a.draws[-20:]
        target_counts = Counter()
        for draw in targets:
            target_counts.update(draw.numbers)

        windows = {
            100: self._window_analyzer(history, 100),
            500: self._window_analyzer(history, 500),
            1000: self._window_analyzer(history, 1000),
        }
        max_input = max(source_weights.values(), default=1)
        scores = {}

        for preset, weights in self.MIXED_PRESETS.items():
            input_w, w100, w500, w1000 = weights
            number_scores = {}

            for number in range(1, 46):
                input_score = source_weights[number] / max_input * 100
                history_scores = []
                for size in (100, 500, 1000):
                    analyzer = windows[size]
                    max_count = max(analyzer.number_counts.values(), default=1)
                    history_scores.append(
                        analyzer.number_counts[number] / max_count * 100
                    )

                number_scores[number] = (
                    input_score * input_w
                    + history_scores[0] * w100
                    + history_scores[1] * w500
                    + history_scores[2] * w1000
                )

            top_numbers = [
                n for n, _ in sorted(
                    number_scores.items(),
                    key=lambda item: (-item[1], item[0]),
                )[:15]
            ]
            scores[preset] = sum(target_counts[n] for n in top_numbers)

        return max(scores, key=scores.get)

    def generate_mixed(
        self,
        source_weights: Counter[int],
        count: int,
        preset: str,
        fixed_numbers: tuple[int, ...] = (),
        excluded_numbers: tuple[int, ...] = (),
        candidate_numbers: tuple[int, ...] = (),
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        """입력번호와 최근 100·500·1000회 데이터를 프리셋별로 다르게 결합합니다."""
        if len(source_weights) < 6:
            raise ValueError("통합데이터추천은 입력번호가 최소 6개 필요합니다.")

        fixed_set = set(fixed_numbers)
        excluded_set = set(excluded_numbers)
        candidate_set = set(candidate_numbers)

        applied_preset = (
            self._auto_mixed_preset(source_weights)
            if preset == "자동최적형"
            else preset
        )
        input_w, w100, w500, w1000 = self.MIXED_PRESETS.get(
            applied_preset,
            self.MIXED_PRESETS["균형형"],
        )

        analyzers = {
            100: self._window_analyzer(self.a.draws, 100),
            500: self._window_analyzer(self.a.draws, 500),
            1000: self._window_analyzer(self.a.draws, 1000),
        }

        max_input = max(source_weights.values(), default=1)
        input_scores = {
            n: source_weights[n] / max_input * 100.0
            for n in range(1, 46)
        }

        window_scores = {}
        for size, analyzer in analyzers.items():
            max_count = max(analyzer.number_counts.values(), default=1)
            window_scores[size] = {
                n: analyzer.number_counts[n] / max_count * 100.0
                for n in range(1, 46)
            }

        combined = {
            n: (
                input_scores[n] * input_w
                + window_scores[100][n] * w100
                + window_scores[500][n] * w500
                + window_scores[1000][n] * w1000
            )
            for n in range(1, 46)
        }

        ranked_input = sorted(range(1, 46), key=lambda n: (-input_scores[n], n))
        ranked100 = sorted(range(1, 46), key=lambda n: (-window_scores[100][n], n))
        ranked500 = sorted(range(1, 46), key=lambda n: (-window_scores[500][n], n))
        ranked1000 = sorted(range(1, 46), key=lambda n: (-window_scores[1000][n], n))
        ranked_combined = sorted(range(1, 46), key=lambda n: (-combined[n], n))

        pool_plan = {
            "입력중심형": (10, 4, 2, 2),
            "최근중심형": (6, 7, 3, 2),
            "균형형": (6, 4, 4, 4),
            "장기형": (4, 2, 5, 7),
        }
        input_n, n100, n500, n1000 = pool_plan.get(
            applied_preset,
            pool_plan["균형형"],
        )

        pool = []
        for number in list(fixed_set | candidate_set):
            if number not in excluded_set and number not in pool:
                pool.append(number)

        for ranked_source, limit in (
            (ranked_input, input_n),
            (ranked100, n100),
            (ranked500, n500),
            (ranked1000, n1000),
        ):
            added = 0
            for number in ranked_source:
                if number in excluded_set or number in pool:
                    continue
                pool.append(number)
                added += 1
                if added >= limit:
                    break

        for number in ranked_combined:
            if number in excluded_set or number in pool:
                continue
            pool.append(number)
            if len(pool) >= 18:
                break

        pool = sorted(pool[:18])

        candidates = []
        for combo in combinations(pool, 6):
            combo_set = set(combo)
            if fixed_set and not fixed_set.issubset(combo_set):
                continue
            if excluded_set & combo_set:
                continue
            if combo in self.a.first_prize or combo in self.a.second_prize:
                continue
            if not 85 <= sum(combo) <= 200:
                continue

            odd = sum(n % 2 for n in combo)
            high = sum(n >= 23 for n in combo)
            if odd not in (2, 3, 4) or high not in (2, 3, 4):
                continue

            metrics = dict(self.metrics(combo, source_weights))
            score100 = self._historical_combo_score(combo, analyzers[100])
            score500 = self._historical_combo_score(combo, analyzers[500])
            score1000 = self._historical_combo_score(combo, analyzers[1000])
            input_score = metrics["input"]
            candidate_hits = len(combo_set & candidate_set)
            candidate_bonus = min(12.0, candidate_hits * 4.0)

            mixed_score = (
                input_score * input_w
                + score100 * w100
                + score500 * w500
                + score1000 * w1000
                + candidate_bonus
            )

            metrics.update({
                "mixed": mixed_score,
                "mixed_preset": applied_preset,
                "mixed_input_weight": input_w,
                "mixed_100_weight": w100,
                "mixed_500_weight": w500,
                "mixed_1000_weight": w1000,
                "mixed_candidate_pool": pool,
                "mixed_top_input": ranked_input[:10],
                "mixed_top_100": ranked100[:10],
                "mixed_top_500": ranked500[:10],
                "mixed_top_1000": ranked1000[:10],
                "score100": score100,
                "score500": score500,
                "score1000": score1000,
                "candidate_hits": candidate_hits,
                "candidate_bonus": candidate_bonus,
                "strategy": f"통합-{applied_preset}",
            })
            candidates.append((mixed_score, combo, metrics))

        candidates.sort(key=lambda row: (-row[0], row[1]))
        return self.select_diverse(candidates, count)

    PATTERN_NAMES = (
        "이월수", "2회전재등장", "단기강세", "장기미출현복귀",
        "끝수흐름", "연속수후보", "동반수확장", "간격수흐름",
    )

    @staticmethod
    def _normalize_counter(
        values: Counter[int] | dict[int, float],
    ) -> dict[int, float]:
        maximum = max(values.values(), default=0)
        if maximum <= 0:
            return {n: 0.0 for n in range(1, 46)}
        return {
            n: float(values.get(n, 0)) / maximum * 100.0
            for n in range(1, 46)
        }

    @staticmethod
    def _number_gaps(draws: list[Draw]) -> dict[int, int]:
        latest_index = len(draws) - 1
        last_seen = {}
        for index, draw in enumerate(draws):
            for number in draw.numbers:
                last_seen[number] = index
        return {
            number: latest_index - last_seen.get(number, -1)
            for number in range(1, 46)
        }

    @classmethod
    def _pattern_number_scores(
        cls,
        draws: list[Draw],
    ) -> tuple[dict[str, dict[int, float]], dict[int, list[str]]]:
        """각 패턴별 1~45 번호 점수와 번호별 추천 근거를 계산합니다."""
        if len(draws) < 10:
            raise ValueError("특이패턴 분석에는 최소 10회 이상의 데이터가 필요합니다.")

        last = draws[-1].numbers
        previous = draws[-2].numbers
        recent10 = draws[-10:]
        recent30 = draws[-30:]
        recent100 = draws[-100:]

        pattern_scores: dict[str, dict[int, float]] = {
            name: {n: 0.0 for n in range(1, 46)}
            for name in cls.PATTERN_NAMES
        }
        reasons: dict[int, list[str]] = defaultdict(list)

        # 1. 이월수: 역대 이월수 평균과 직전 회차 번호
        rollover_counts = Counter()
        for before, after in zip(draws[:-1], draws[1:]):
            rollover_counts[len(set(before.numbers) & set(after.numbers))] += 1
        total_transitions = max(1, sum(rollover_counts.values()))
        rollover_probability = 1.0 - rollover_counts[0] / total_transitions
        rollover_base = 72.0 + rollover_probability * 25.0
        for number in last:
            pattern_scores["이월수"][number] = rollover_base
            reasons[number].append("직전회차 이월수 후보")

        # 2. 2회 전 재등장: 2회 전에는 있었지만 직전에는 없던 번호
        for number in set(previous) - set(last):
            pattern_scores["2회전재등장"][number] = 82.0
            reasons[number].append("2회 전 번호 재등장 후보")

        # 3. 단기강세: 최근 10회와 30회 빈도를 혼합
        count10 = Counter(n for draw in recent10 for n in draw.numbers)
        count30 = Counter(n for draw in recent30 for n in draw.numbers)
        norm10 = cls._normalize_counter(count10)
        norm30 = cls._normalize_counter(count30)
        for number in range(1, 46):
            score = norm10[number] * 0.65 + norm30[number] * 0.35
            pattern_scores["단기강세"][number] = score
            if score >= 72:
                reasons[number].append("최근 10·30회 강세")

        # 4. 장기 미출현 복귀
        gaps = cls._number_gaps(draws)
        sorted_gaps = sorted(gaps.values())
        q70 = sorted_gaps[int(len(sorted_gaps) * 0.70)]
        max_gap = max(sorted_gaps, default=1)
        for number, gap in gaps.items():
            score = min(100.0, gap / max(1, max_gap) * 100.0)
            pattern_scores["장기미출현복귀"][number] = score
            if gap >= q70:
                reasons[number].append(f"{gap}회 미출현 복귀 후보")

        # 5. 끝수 흐름
        ending10 = Counter(n % 10 for draw in recent10 for n in draw.numbers)
        ending30 = Counter(n % 10 for draw in recent30 for n in draw.numbers)
        max_ending = max(
            (ending10[e] * 0.7 + ending30[e] * 0.3 for e in range(10)),
            default=1,
        )
        for number in range(1, 46):
            ending = number % 10
            raw = ending10[ending] * 0.7 + ending30[ending] * 0.3
            score = raw / max(1, max_ending) * 100.0
            pattern_scores["끝수흐름"][number] = score
            if score >= 78:
                reasons[number].append(f"끝수 {ending} 흐름 강세")

        # 6. 연속수 후보: 직전 번호의 앞뒤 번호
        for number in last:
            for candidate in (number - 1, number + 1):
                if 1 <= candidate <= 45 and candidate not in last:
                    pattern_scores["연속수후보"][candidate] = max(
                        pattern_scores["연속수후보"][candidate],
                        88.0,
                    )
                    reasons[candidate].append(f"{number}번 인접 연속수 후보")

        # 7. 동반수 확장: 직전 번호들과 최근100회에 자주 함께 나온 번호
        recent_pairs = Counter(
            pair for draw in recent100 for pair in combinations(draw.numbers, 2)
        )
        partner_raw = Counter()
        for number in last:
            for candidate in range(1, 46):
                if candidate == number or candidate in last:
                    continue
                pair = tuple(sorted((number, candidate)))
                partner_raw[candidate] += recent_pairs[pair]
        partner_norm = cls._normalize_counter(partner_raw)
        for number, score in partner_norm.items():
            pattern_scores["동반수확장"][number] = score
            if score >= 72:
                reasons[number].append("직전번호 동반수 확장")

        # 8. 간격수 흐름: 최근 당첨 조합에서 자주 나온 번호 간 차이를 직전번호에 적용
        gap_counts = Counter()
        for draw in recent30:
            nums = draw.numbers
            for a, b in combinations(nums, 2):
                gap = b - a
                if 1 <= gap <= 15:
                    gap_counts[gap] += 1
        common_gaps = [gap for gap, _ in gap_counts.most_common(5)]
        for source in last:
            for rank, gap in enumerate(common_gaps):
                score = 92.0 - rank * 8.0
                for candidate in (source - gap, source + gap):
                    if 1 <= candidate <= 45 and candidate not in last:
                        pattern_scores["간격수흐름"][candidate] = max(
                            pattern_scores["간격수흐름"][candidate],
                            score,
                        )
                        if score >= 76:
                            reasons[candidate].append(f"간격 {gap} 흐름 후보")

        return pattern_scores, reasons

    @classmethod
    def pattern_reliability(
        cls,
        draws: list[Draw],
        test_rounds: int = 60,
    ) -> dict[str, float]:
        """최근 과거 회차에서 패턴별 TOP10의 평균 적중도를 계산합니다."""
        reliability = {name: 50.0 for name in cls.PATTERN_NAMES}
        if len(draws) < 80:
            return reliability

        hit_totals = Counter()
        tested = 0
        start = max(20, len(draws) - test_rounds)
        for target_index in range(start, len(draws)):
            history = draws[:target_index]
            target = set(draws[target_index].numbers)
            scores, _ = cls._pattern_number_scores(history)
            for name, number_scores in scores.items():
                top10 = {
                    n for n, _ in sorted(
                        number_scores.items(),
                        key=lambda item: (-item[1], item[0]),
                    )[:10]
                }
                hit_totals[name] += len(top10 & target)
            tested += 1

        if tested:
            for name in cls.PATTERN_NAMES:
                average_hits = hit_totals[name] / tested
                # TOP10 무작위 기대 적중은 약 1.33개. 이를 기준으로 35~100점 환산.
                reliability[name] = round(
                    max(35.0, min(100.0, 50.0 + (average_hits - 1.33) * 28.0)),
                    1,
                )
        return reliability

    def pattern_board(
        self,
    ) -> tuple[list[dict[str, object]], dict[str, float]]:
        scores, reasons = self._pattern_number_scores(self.a.draws)
        reliability = self.pattern_reliability(self.a.draws)

        rows = []
        for number in range(1, 46):
            votes = []
            weighted_score = 0.0
            weight_sum = 0.0
            for name in self.PATTERN_NAMES:
                score = scores[name][number]
                rel = reliability[name]
                weighted_score += score * rel
                weight_sum += rel
                if score >= 68:
                    votes.append(name)

            final_score = weighted_score / max(1.0, weight_sum)
            rows.append({
                "number": number,
                "score": round(final_score, 1),
                "votes": len(votes),
                "patterns": votes,
                "reasons": reasons.get(number, []),
            })

        rows.sort(
            key=lambda row: (
                -int(row["votes"]),
                -float(row["score"]),
                int(row["number"]),
            )
        )
        return rows, reliability

    def pattern_briefing(self) -> str:
        board, reliability = self.pattern_board()
        top = board[:12]
        exclusions = sorted(
            board,
            key=lambda row: (
                int(row["votes"]),
                float(row["score"]),
                int(row["number"]),
            ),
        )[:8]

        latest = self.a.draws[-1]
        recent10 = self.a.draws[-10:]
        average_sum = sum(sum(d.numbers) for d in recent10) / len(recent10)
        average_odd = sum(
            sum(n % 2 for n in d.numbers) for d in recent10
        ) / len(recent10)
        strongest = sorted(
            reliability.items(),
            key=lambda item: (-item[1], item[0]),
        )[:3]

        top_text = " · ".join(
            f"{row['number']}({row['votes']}표)"
            for row in top
        )
        excluded_text = " · ".join(str(row["number"]) for row in exclusions)
        pattern_text = " / ".join(
            f"{name} {score:.1f}점" for name, score in strongest
        )
        return (
            f"이번 주 특이패턴 브리핑\n"
            f"최신 기준 회차: {latest.round_no}회\n"
            f"최근 10회 평균 합계: {average_sum:.1f} / 평균 홀수: {average_odd:.1f}개\n"
            f"검증점수 상위 패턴: {pattern_text}\n\n"
            f"패턴투표 핵심번호 TOP12\n{top_text}\n\n"
            f"AI 제외 검토번호 8개\n{excluded_text}\n\n"
            f"※ 패턴투표는 과거 통계를 이용한 분석이며 당첨을 보장하지 않습니다."
        )

    def historical_similar_draws(
        self,
        combo: tuple[int, ...],
        top_n: int = 5,
    ) -> list[tuple[int, float, tuple[int, ...]]]:
        """합계·홀짝·구간·끝수·연속수 특성이 비슷한 과거 회차를 찾습니다."""
        def signature(numbers: tuple[int, ...]) -> tuple:
            total = sum(numbers)
            odd = sum(n % 2 for n in numbers)
            low = sum(n <= 22 for n in numbers)
            endings = len({n % 10 for n in numbers})
            consecutive = self.consecutive_pairs(numbers)
            zones = (
                sum(1 <= n <= 10 for n in numbers),
                sum(11 <= n <= 20 for n in numbers),
                sum(21 <= n <= 30 for n in numbers),
                sum(31 <= n <= 40 for n in numbers),
                sum(41 <= n <= 45 for n in numbers),
            )
            return total, odd, low, endings, consecutive, zones

        target = signature(combo)
        rows = []
        for draw in self.a.draws:
            sig = signature(draw.numbers)
            distance = (
                abs(target[0] - sig[0]) / 25.0
                + abs(target[1] - sig[1]) * 0.8
                + abs(target[2] - sig[2]) * 0.6
                + abs(target[3] - sig[3]) * 0.4
                + abs(target[4] - sig[4]) * 0.8
                + sum(abs(a - b) for a, b in zip(target[5], sig[5])) * 0.35
            )
            similarity = max(0.0, 100.0 - distance * 12.0)
            rows.append((draw.round_no, round(similarity, 1), draw.numbers))
        rows.sort(key=lambda row: (-row[1], -row[0]))
        return rows[:top_n]

    def generate_pattern(
        self,
        count: int,
        mode: str = "자동종합",
        fixed_numbers: tuple[int, ...] = (),
        excluded_numbers: tuple[int, ...] = (),
        candidate_numbers: tuple[int, ...] = (),
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        board, reliability = self.pattern_board()
        board_map = {int(row["number"]): row for row in board}
        fixed_set = set(fixed_numbers)
        excluded_set = set(excluded_numbers)
        candidate_set = set(candidate_numbers)

        if mode != "자동종합" and mode in self.PATTERN_NAMES:
            pattern_scores, reasons = self._pattern_number_scores(self.a.draws)
            ranked_numbers = sorted(
                range(1, 46),
                key=lambda n: (-pattern_scores[mode][n], n),
            )
        else:
            ranked_numbers = [int(row["number"]) for row in board]
            _, reasons = self._pattern_number_scores(self.a.draws)

        selected = []
        for number in list(fixed_set | candidate_set) + ranked_numbers:
            if number in excluded_set or number in selected:
                continue
            selected.append(number)
            if len(selected) >= 20:
                break
        pool = sorted(selected)

        candidates = []
        for combo in combinations(pool, 6):
            combo_set = set(combo)
            if fixed_set and not fixed_set.issubset(combo_set):
                continue
            if excluded_set & combo_set:
                continue
            if combo in self.a.first_prize or combo in self.a.second_prize:
                continue

            odd = sum(n % 2 for n in combo)
            high = sum(n >= 23 for n in combo)
            total = sum(combo)
            if odd not in (2, 3, 4) or high not in (2, 3, 4):
                continue
            if not 85 <= total <= 200:
                continue

            values = [float(board_map[n]["score"]) for n in combo]
            votes = [int(board_map[n]["votes"]) for n in combo]
            pattern_score = sum(values) / 6.0 + sum(votes) * 1.7

            pair_values = [
                self.a.recent_pair_counts[pair]
                for pair in combinations(combo, 2)
            ]
            pair_bonus = min(12.0, sum(sorted(pair_values, reverse=True)[:4]) / 8.0)
            candidate_bonus = min(12.0, len(combo_set & candidate_set) * 4.0)
            final_score = pattern_score + pair_bonus + candidate_bonus

            metrics = dict(self.metrics(combo, Counter({n: 1 for n in pool})))
            combo_patterns = Counter()
            combo_reasons = []
            for number in combo:
                for pattern in board_map[number]["patterns"]:
                    combo_patterns[pattern] += 1
                combo_reasons.extend(reasons.get(number, []))

            leading_patterns = [
                name for name, _ in combo_patterns.most_common(3)
            ]
            metrics.update({
                "pattern": final_score,
                "pattern_score": pattern_score,
                "pattern_votes": sum(votes),
                "pattern_mode": mode,
                "pattern_names": leading_patterns,
                "pattern_reasons": list(dict.fromkeys(combo_reasons))[:5],
                "pattern_reliability": reliability,
                "strategy": f"특이패턴-{mode}",
            })
            candidates.append((final_score, combo, metrics))

        candidates.sort(key=lambda row: (-row[0], row[1]))
        return self.select_diverse(candidates, count)

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

    def v27_candidate_pool(self, size: int = 25) -> tuple[list[int], dict[int, float]]:
        """v27 페어다양성100 방식: 번호점수와 페어 연결성·구간 다양성을 함께 사용합니다."""
        base = self.self_number_scores()
        max_pair = max(self.a.pair_counts.values(), default=1)
        pair_strength: dict[int, float] = {}
        for number in range(1, 46):
            values = sorted(
                (self.a.pair_counts[tuple(sorted((number, other)))] for other in range(1, 46) if other != number),
                reverse=True,
            )[:8]
            pair_strength[number] = (sum(values) / max(1, len(values))) / max_pair * 100.0

        score = {n: base[n] * 0.64 + pair_strength[n] * 0.36 for n in range(1, 46)}
        ranked = sorted(range(1, 46), key=lambda n: (-score[n], n))
        selected: list[int] = []
        zones = [(1, 10), (11, 20), (21, 30), (31, 40), (41, 45)]

        # 각 구간의 핵심번호를 먼저 확보합니다.
        for lo, hi in zones:
            selected.extend([n for n in ranked if lo <= n <= hi][:2])

        # 이미 선택된 번호와의 페어 연결성과 과도한 유사성을 동시에 평가합니다.
        while len(selected) < size:
            best = None
            best_value = float('-inf')
            for n in ranked:
                if n in selected:
                    continue
                links = [self.a.pair_counts[tuple(sorted((n, old)))] for old in selected]
                link_score = sum(sorted(links, reverse=True)[:5]) / max(1, max_pair * 5) * 100.0
                same_end_penalty = sum(1 for old in selected if old % 10 == n % 10) * 1.8
                neighbor_penalty = sum(1 for old in selected if abs(old - n) <= 1) * 1.2
                value = score[n] * 0.72 + link_score * 0.28 - same_end_penalty - neighbor_penalty
                if value > best_value:
                    best_value = value
                    best = n
            if best is None:
                break
            selected.append(best)
        return sorted(selected[:size]), score

    def generate_self_v27_v40(
        self,
        count: int = 100,
        mode: str = "자동종합",
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        """자체추천 전용: v27이 후보를 만들고 v40이 최종 조합을 선별합니다."""
        pool, number_scores = self.v27_candidate_pool(25)
        source_weights = Counter({n: max(1, round(number_scores[n])) for n in pool})
        candidates = []
        for combo in combinations(pool, 6):
            total = sum(combo)
            odd = sum(n % 2 for n in combo)
            high = sum(n >= 23 for n in combo)
            if not 85 <= total <= 200 or odd not in (2, 3, 4) or high not in (2, 3, 4):
                continue
            if self.consecutive_pairs(combo) > 2:
                continue
            if combo in self.a.first_prize or combo in self.a.second_prize:
                continue
            metrics = dict(self.metrics(combo, source_weights))
            candidate_score = sum(number_scores[n] for n in combo) / 6.0
            pair_score = metrics.get('pair', 0.0)
            triple_score = metrics.get('triple', 0.0)
            final_score = candidate_score * 0.52 + pair_score * 0.20 + triple_score * 0.10 + metrics.get('structure', 0.0) * 0.18
            metrics.update({
                'self': final_score,
                'candidate_engine': 'v27 페어다양성100',
                'combination_engine': 'v40 Elite Survival',
                'self_candidate_pool': pool,
                'self_pattern_mode': mode,
                'self_pattern_mix': 0.0,
                'self_pattern_top10': [],
                'filter_mode': 'v27 후보25 → v40 조합',
            })
            candidates.append((final_score, combo, metrics))
        candidates.sort(key=lambda row: (-row[0], row[1]))
        return self.select_diverse(candidates, count)

    def generate_self(
        self,
        count: int = 100,
        mode: str = "자동종합",
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        """역대 전체 데이터와 선택한 특이패턴을 결합해 자체추천합니다."""
        base_scores = self.self_number_scores()
        pattern_scores, pattern_reasons = self._pattern_number_scores(self.a.draws)
        reliability = self.pattern_reliability(self.a.draws)

        # 선택 패턴이 실제 번호점수와 후보 풀에 반영되도록 구성합니다.
        selected_scores: dict[int, float] = {}
        if mode == "자동종합":
            for number in range(1, 46):
                weighted = 0.0
                weight_sum = 0.0
                for pattern_name in self.PATTERN_NAMES:
                    rel = reliability.get(pattern_name, 50.0)
                    weighted += pattern_scores[pattern_name][number] * rel
                    weight_sum += rel
                selected_scores[number] = weighted / max(1.0, weight_sum)
            pattern_mix = 0.48
        elif mode in pattern_scores:
            selected_scores = dict(pattern_scores[mode])
            pattern_mix = 0.68
        else:
            selected_scores = {number: 0.0 for number in range(1, 46)}
            pattern_mix = 0.0

        number_scores = {
            number: (
                base_scores[number] * (1.0 - pattern_mix)
                + selected_scores[number] * pattern_mix
            )
            for number in range(1, 46)
        }

        ranked = sorted(number_scores, key=lambda n: (-number_scores[n], n))
        pattern_ranked = sorted(
            selected_scores, key=lambda n: (-selected_scores[n], n)
        )

        # 선택한 패턴의 상위 번호를 후보번호 풀에 직접 포함합니다.
        pool = set(ranked[:18])
        pool.update(pattern_ranked[:10])
        for low, high in [(1, 10), (11, 20), (21, 30), (31, 40), (41, 45)]:
            pool.update([n for n in ranked if low <= n <= high][:1])
            pool.update([n for n in pattern_ranked if low <= n <= high][:1])

        pool = sorted(
            pool,
            key=lambda n: (
                -(number_scores[n] * 0.55 + selected_scores[n] * 0.45),
                n,
            ),
        )[:28]
        pool = sorted(pool)

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
            selected_pattern_score = sum(selected_scores[n] for n in combo) / 6.0

            # 선택 패턴은 번호 후보 선정뿐 아니라 최종 조합 순위에도 직접 반영합니다.
            self_score = (
                own_number_score * 0.34
                + metrics["composite"] * 0.33
                + selected_pattern_score * 0.33
            )
            metrics = dict(metrics)
            metrics["self"] = self_score
            metrics["self_pattern_mode"] = mode
            metrics["self_pattern_score"] = selected_pattern_score
            metrics["pattern_names"] = [mode] if mode != "자동종합" else ["자동종합"]
            metrics["pattern_reasons"] = list(dict.fromkeys(
                reason
                for number in combo
                for reason in pattern_reasons.get(number, [])
            ))[:5]
            metrics["self_pattern_mix"] = pattern_mix
            metrics["self_candidate_pool"] = pool
            metrics["self_pattern_top10"] = pattern_ranked[:10]
            candidates.append((self_score, combo, metrics))

        candidates.sort(key=lambda x: (-x[0], x[1]))
        return self.select_diverse(candidates, count)

    def select_pool(
        self,
        source_weights: Counter[int],
        category: str,
        limit: int = 24,
    ) -> list[int]:
        """입력번호가 많아도 계산이 멈추지 않도록 카테고리별 핵심 번호를 선별합니다."""
        numbers = sorted(source_weights)
        if len(numbers) <= limit:
            return numbers

        max_input = max(source_weights.values(), default=1)

        pair_centrality = {}
        triple_centrality = {}
        for n in numbers:
            pair_centrality[n] = sum(
                self.a.pair_counts[tuple(sorted((n, other)))]
                for other in numbers
                if other != n
            )
            triple_centrality[n] = sum(
                count
                for triple, count in self.a.triple_counts.items()
                if n in triple
            )

        max_pair_c = max(pair_centrality.values(), default=1)
        max_triple_c = max(triple_centrality.values(), default=1)
        max_recent = max(
            (self.a.recent_number_counts[n] for n in numbers),
            default=1,
        )

        scores = {}
        for n in numbers:
            input_score = source_weights[n] / max_input * 100
            pair_score = pair_centrality[n] / max_pair_c * 100
            triple_score = triple_centrality[n] / max_triple_c * 100
            recent_score = self.a.recent_number_counts[n] / max_recent * 100

            if category == "나온횟수":
                score = input_score * 0.75 + pair_score * 0.10 + recent_score * 0.15
            elif category == "동반수":
                score = pair_score * 0.75 + input_score * 0.15 + recent_score * 0.10
            elif category == "트리플":
                score = triple_score * 0.75 + pair_score * 0.15 + input_score * 0.10
            elif category == "최근패턴":
                score = recent_score * 0.75 + pair_score * 0.15 + input_score * 0.10
            else:
                score = (
                    input_score * 0.30
                    + pair_score * 0.25
                    + triple_score * 0.20
                    + recent_score * 0.25
                )
            scores[n] = score

        ranked = sorted(numbers, key=lambda n: (-scores[n], n))

        # 각 번호 구간에서 최소 2개씩 포함해 특정 구간 쏠림을 방지
        selected = set(ranked[: max(14, limit - 10)])
        for low, high in [(1, 10), (11, 20), (21, 30), (31, 40), (41, 45)]:
            zone = [n for n in ranked if low <= n <= high][:2]
            selected.update(zone)

        final = sorted(selected, key=lambda n: (-scores[n], n))[:limit]
        return sorted(final)

    @staticmethod
    def filter_mode(input_count: int) -> tuple[str, str]:
        """입력 번호 개수에 따라 자동 필터 강도를 결정합니다."""
        if input_count <= 14:
            return "기본", "10~14개 입력: 결과 확보를 우선하는 완화 필터"
        if input_count <= 24:
            return "고급", "15~24개 입력: 균형과 통계를 함께 보는 고급 필터"
        return "정밀", "25개 이상 입력: 후보가 많아 더 엄격한 정밀 필터"

    def passes_filter(
        self,
        combo: tuple[int, ...],
        mode: str,
        sum_min: int,
        sum_max: int,
        allow_consecutive: bool,
    ) -> bool:
        total = sum(combo)
        if not sum_min <= total <= sum_max:
            return False

        odd = sum(n % 2 for n in combo)
        high = sum(n >= 23 for n in combo)
        consecutive = self.consecutive_pairs(combo)

        if not allow_consecutive and consecutive > 0:
            return False

        if mode == "기본":
            if odd not in (1, 2, 3, 4, 5):
                return False
            if high not in (0, 1, 2, 3, 4, 5, 6):
                return False
            if consecutive > 4:
                return False
            return True

        if mode == "고급":
            if odd not in (2, 3, 4):
                return False
            if high not in (2, 3, 4):
                return False
            if consecutive > 2:
                return False
            return True

        if odd not in (2, 3, 4):
            return False
        if high not in (2, 3, 4):
            return False
        if consecutive > 1:
            return False
        if total < 85 or total > 195:
            return False
        return True

    def relaxed_fallback(
        self,
        pool: list[int],
        source_weights: Counter[int],
        category: str,
        count: int,
        sum_min: int,
        sum_max: int,
        fixed_numbers: tuple[int, ...] = (),
        excluded_numbers: tuple[int, ...] = (),
        candidate_numbers: tuple[int, ...] = (),
        strategy: str = "균형형",
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        """필터 결과가 부족할 때 점수순으로 자동 보충합니다."""
        candidates = []
        for combo in combinations(pool, 6):
            combo_set = set(combo)
            if fixed_numbers and not set(fixed_numbers).issubset(combo_set):
                continue
            if excluded_numbers and combo_set.intersection(excluded_numbers):
                continue
            total = sum(combo)
            if not sum_min <= total <= sum_max:
                continue
            if combo in self.a.first_prize or combo in self.a.second_prize:
                continue

            metrics = dict(self.metrics(combo, source_weights))
            candidate_hits = len(combo_set.intersection(candidate_numbers))
            candidate_bonus = min(12.0, candidate_hits * 4.0)
            metrics["candidate_hits"] = candidate_hits
            metrics["candidate_bonus"] = candidate_bonus
            base_score = (
                self.strategy_score(metrics, strategy)
                if category == "추천조합"
                else self.category_score(metrics, category)
            )
            metrics["strategy"] = strategy
            score = base_score + candidate_bonus
            candidates.append((score, combo, metrics))

        candidates.sort(key=lambda x: (-x[0], x[1]))
        return candidates[:count]

    def recommendation_reason(
        self,
        metrics: dict[str, float],
        fixed_numbers: tuple[int, ...] = (),
        candidate_numbers: tuple[int, ...] = (),
        combo: tuple[int, ...] = (),
    ) -> str:
        """추천 근거를 한 줄로 요약합니다."""
        labels = [
            ("나온횟수 강함", metrics.get("input", 0.0)),
            ("동반수 강함", metrics.get("pair", 0.0)),
            ("트리플 강함", metrics.get("triple", 0.0)),
            ("최근패턴 우수", metrics.get("recent", 0.0)),
            ("조합 균형 우수", metrics.get("structure", 0.0)),
        ]
        labels.sort(key=lambda item: (-item[1], item[0]))
        selected = [name for name, score in labels[:2] if score >= 35]
        if not selected:
            selected = [labels[0][0]]

        if fixed_numbers:
            selected.append("필수번호 " + ",".join(map(str, fixed_numbers)))

        included_candidates = sorted(set(combo) & set(candidate_numbers))
        if included_candidates:
            selected.append(
                "후보번호 포함 " + ",".join(map(str, included_candidates))
            )

        return " / ".join(selected)

    def generate(
        self,
        source_weights: Counter[int],
        count: int,
        sum_min: int,
        sum_max: int,
        allow_consecutive: bool,
        category: str,
        fixed_numbers: tuple[int, ...] = (),
        excluded_numbers: tuple[int, ...] = (),
        candidate_numbers: tuple[int, ...] = (),
        strategy: str = "균형형",
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        input_count = len(source_weights)
        mode, _ = self.filter_mode(input_count)

        fixed_numbers = tuple(sorted(set(fixed_numbers)))
        excluded_numbers = tuple(sorted(set(excluded_numbers)))
        candidate_numbers = tuple(sorted(set(candidate_numbers)))

        fixed_set = set(fixed_numbers)
        excluded_set = set(excluded_numbers)
        candidate_set = set(candidate_numbers)

        if fixed_set & excluded_set:
            raise ValueError("필수번호와 제외번호가 중복됩니다.")
        if fixed_set & candidate_set:
            raise ValueError("필수번호와 후보번호가 중복됩니다.")
        if excluded_set & candidate_set:
            raise ValueError("제외번호와 후보번호가 중복됩니다.")

        missing_fixed = [n for n in fixed_numbers if n not in source_weights]
        if missing_fixed:
            raise ValueError(
                "필수번호는 번호 입력란에도 포함되어야 합니다: "
                + ", ".join(map(str, missing_fixed))
            )

        pool_limit = 24 if input_count >= 25 else input_count
        pool = self.select_pool(source_weights, category, limit=pool_limit)
        pool = sorted((set(pool) | fixed_set | candidate_set) - excluded_set)

        if len(pool) < 6:
            raise ValueError("고유 번호가 최소 6개 필요합니다.")

        candidates = []
        for combo in combinations(pool, 6):
            combo_set = set(combo)
            if fixed_numbers and not set(fixed_numbers).issubset(combo_set):
                continue
            if excluded_numbers and combo_set.intersection(excluded_numbers):
                continue
            if not self.passes_filter(
                combo, mode, sum_min, sum_max, allow_consecutive
            ):
                continue
            if combo in self.a.first_prize or combo in self.a.second_prize:
                continue

            metrics = dict(self.metrics(combo, source_weights))
            metrics["filter_mode"] = mode
            candidate_hits = len(combo_set.intersection(candidate_numbers))
            candidate_bonus = min(12.0, candidate_hits * 4.0)
            metrics["candidate_hits"] = candidate_hits
            metrics["candidate_bonus"] = candidate_bonus
            base_score = (
                self.strategy_score(metrics, strategy)
                if category == "추천조합"
                else self.category_score(metrics, category)
            )
            metrics["strategy"] = strategy
            score = base_score + candidate_bonus
            candidates.append((score, combo, metrics))

        candidates.sort(key=lambda x: (-x[0], x[1]))

        if len(candidates) < count:
            existing = {combo for _, combo, _ in candidates}
            fallback = self.relaxed_fallback(
                pool, source_weights, category, count, sum_min, sum_max,
                fixed_numbers=fixed_numbers,
                excluded_numbers=excluded_numbers,
                candidate_numbers=candidate_numbers,
                strategy=strategy,
            )
            for score, combo, metrics in fallback:
                if combo in existing:
                    continue
                metrics = dict(metrics)
                metrics["filter_mode"] = f"{mode}→자동완화"
                candidates.append((score, combo, metrics))
                existing.add(combo)
                if len(candidates) >= count:
                    break

        candidates.sort(key=lambda x: (-x[0], x[1]))
        return self.select_diverse(candidates, count)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.analyzer = LottoAnalyzer()
        self.photo_paths: list[str] = []
        self.recommendations: list[tuple[float, tuple[int, ...], dict[str, float]]] = []
        self.ocr_cache: dict[tuple[str, int, int], list[int]] = {}
        self.pattern_cache: dict[str, object] = {}
        self.suspend_auto_recommend = False
        self.ocr_worker: OCRWorker | None = None

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
        self.statusBar().showMessage("이전 작업상태를 확인하고 있습니다...")
        self.restore_session()

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

        names = [
            "번호 입력", "추천조합", "나온횟수", "동반수", "트리플",
            "최근패턴", "통합데이터추천", "성과최적추천", "AI진화추천", "AI연구소추천", "특이패턴추천", "자체추천"
        ]
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

        data_menu = self.menuBar().addMenu("데이터 관리")
        manual_action = QAction("수동 회차 추가", self)
        manual_action.triggered.connect(self.manual_add_draw)
        data_menu.addAction(manual_action)

        latest_action = QAction("최신 회차 자동 확인", self)
        latest_action.triggered.connect(self.check_latest_draw)
        data_menu.addAction(latest_action)

        performance_action = QAction("최적화·항목별 결과", self)
        performance_action.triggered.connect(self.show_performance_report)
        data_menu.addAction(performance_action)

        save_action = QAction("현재 상태 저장", self)
        save_action.triggered.connect(self.save_session)
        data_menu.addAction(save_action)
        check_action = QAction("내부 자동검증 결과", self)
        check_action.triggered.connect(self.show_self_check_report)
        data_menu.addAction(check_action)

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

        self.excel_status = QLabel(
            "역대 Excel이 아직 등록되지 않았습니다. 왼쪽 아래 '역대 Excel 불러오기'를 누르세요."
        )
        self.excel_status.setObjectName("card")
        self.excel_status.setWordWrap(True)
        lay.addWidget(self.excel_status)

        self.excel_progress = QProgressBar()
        self.excel_progress.setRange(0, 100)
        self.excel_progress.setValue(0)
        lay.addWidget(self.excel_progress)

        grid = QGridLayout()
        left = QFrame()
        left.setObjectName("card")
        ll = QVBoxLayout(left)
        ll.addWidget(QLabel("사진 파일 등록"))
        self.photo_list = QListWidget()
        ll.addWidget(self.photo_list)
        row = QHBoxLayout()
        self.photo_add_button = QPushButton("사진 추가·자동 인식")
        self.photo_add_button.clicked.connect(self.add_photos)
        self.photo_delete_button = QPushButton("선택 삭제")
        self.photo_delete_button.clicked.connect(self.delete_photo)
        self.photo_rerun_button = QPushButton("선택 사진 다시 인식")
        self.photo_rerun_button.clicked.connect(self.rerun_selected_photo_ocr)
        row.addWidget(self.photo_add_button)
        row.addWidget(self.photo_delete_button)
        ll.addLayout(row)
        ll.addWidget(self.photo_rerun_button)

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

        fixed_label = QLabel(
            "필수번호 — 자체추천을 제외한 모든 추천 조합에 반드시 포함됩니다."
        )
        fixed_label.setWordWrap(True)
        rl.addWidget(fixed_label)

        self.fixed_input = QLineEdit()
        self.fixed_input.setPlaceholderText("예: 3 6 또는 3, 6, 24")
        rl.addWidget(self.fixed_input)

        excluded_label = QLabel(
            "제외번호 — 자체추천을 제외한 모든 추천 조합에서 제거됩니다."
        )
        excluded_label.setWordWrap(True)
        rl.addWidget(excluded_label)

        self.excluded_input = QLineEdit()
        self.excluded_input.setPlaceholderText("예: 18 29")
        rl.addWidget(self.excluded_input)

        candidate_label = QLabel(
            "후보번호 — 포함 시 가산점을 받지만 필수는 아닙니다."
        )
        candidate_label.setWordWrap(True)
        rl.addWidget(candidate_label)

        self.candidate_input = QLineEdit()
        self.candidate_input.setPlaceholderText("예: 7 14 33")
        rl.addWidget(self.candidate_input)

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
            "추천 100조합 · 역대 1등·2등 동일 조합 제외\n"
            "특이패턴추천은 이월수·재등장·강세·미출현·끝수·연속수·동반수·간격수의 번호를 투표식으로 종합합니다.\n"
            "자체추천은 v27 페어다양성100이 후보번호를 만들고 v40 Elite Survival이 조합합니다. 그 외 모든 추천 항목도 최종 조합 선별에는 v40 엔진을 사용합니다."
        )
        guide.setObjectName("card")
        guide.setWordWrap(True)
        lay.addWidget(guide)

        self.rec_category = QComboBox()
        self.rec_category.addItems(
            [
                "추천조합", "나온횟수", "동반수", "트리플",
                "최근패턴", "통합데이터추천", "성과최적추천",
                "AI진화추천", "AI연구소추천", "AI후보번호연구", "AI이중후보추천",
                "특이패턴추천", "자체추천"
            ]
        )
        self.rec_category.currentTextChanged.connect(self.generate_recommendations)
        lay.addWidget(self.rec_category)

        strategy_row = QHBoxLayout()
        strategy_label = QLabel("추천 전략")
        self.strategy_combo = QComboBox()
        self.strategy_combo.addItems(
            ["균형형", "출현형", "동반수형", "트리플형", "최근형", "AI Ultimate"]
        )
        self.strategy_combo.currentTextChanged.connect(self.generate_recommendations)
        strategy_row.addWidget(strategy_label)
        strategy_row.addWidget(self.strategy_combo)

        self.strategy_battle_btn = QPushButton("전략 배틀 백테스트")
        self.strategy_battle_btn.clicked.connect(self.run_strategy_battle)
        strategy_row.addWidget(self.strategy_battle_btn)
        lay.addLayout(strategy_row)

        mixed_row = QHBoxLayout()
        mixed_label = QLabel("통합데이터 비율")
        self.mixed_preset_combo = QComboBox()
        self.mixed_preset_combo.addItems(
            ["입력중심형", "최근중심형", "균형형", "장기형", "자동최적형"]
        )
        self.mixed_preset_combo.currentTextChanged.connect(
            self.on_mixed_preset_changed
        )
        mixed_row.addWidget(mixed_label)
        mixed_row.addWidget(self.mixed_preset_combo)
        mixed_help = QLabel(
            "입력번호 + 최근 100회 + 500회 + 1000회 데이터를 섞어 새 조합을 만듭니다."
        )
        mixed_help.setWordWrap(True)
        mixed_row.addWidget(mixed_help, 1)
        lay.addLayout(mixed_row)

        pattern_row = QHBoxLayout()
        pattern_label = QLabel("특이패턴")
        self.pattern_mode_combo = QComboBox()
        self.pattern_mode_combo.addItems(
            [
                "자동종합", "이월수", "2회전재등장", "단기강세",
                "장기미출현복귀", "끝수흐름", "연속수후보",
                "동반수확장", "간격수흐름",
            ]
        )
        self.pattern_mode_combo.currentTextChanged.connect(
            self.on_pattern_mode_changed
        )
        pattern_row.addWidget(pattern_label)
        pattern_row.addWidget(self.pattern_mode_combo)

        self.pattern_brief_btn = QPushButton("이번 주 패턴 브리핑")
        self.pattern_brief_btn.clicked.connect(self.show_pattern_briefing)
        pattern_row.addWidget(self.pattern_brief_btn)
        self.performance_report_btn = QPushButton("최적화·항목별 결과")
        self.performance_report_btn.clicked.connect(self.show_performance_report)
        pattern_row.addWidget(self.performance_report_btn)
        self.evolution_report_btn = QPushButton("AI 성장결과")
        self.evolution_report_btn.clicked.connect(self.show_evolution_report)
        pattern_row.addWidget(self.evolution_report_btn)
        self.lab_report_btn = QPushButton("AI 연구소 결과")
        self.lab_report_btn.clicked.connect(self.show_lab_report)
        pattern_row.addWidget(self.lab_report_btn)
        self.candidate_report_btn = QPushButton("후보15 연구결과")
        self.candidate_report_btn.clicked.connect(self.show_candidate_lab_report)
        pattern_row.addWidget(self.candidate_report_btn)
        self.shrink_report_btn = QPushButton("후보축소 검증")
        self.shrink_report_btn.clicked.connect(self.show_candidate_shrink_report)
        pattern_row.addWidget(self.shrink_report_btn)
        self.dual_candidate_report_btn = QPushButton("이중후보 검증")
        self.dual_candidate_report_btn.clicked.connect(self.show_dual_candidate_report)
        pattern_row.addWidget(self.dual_candidate_report_btn)
        self.candidate22_report_btn = QPushButton("후보22 축소검증")
        self.candidate22_report_btn.clicked.connect(self.show_candidate22_report)
        pattern_row.addWidget(self.candidate22_report_btn)

        self.round_search_input = QLineEdit()
        self.round_search_input.setPlaceholderText("회차/번호 검색: 33 37 40")
        pattern_row.addWidget(self.round_search_input, 1)
        self.round_search_btn = QPushButton("회차 검색")
        self.round_search_btn.clicked.connect(self.search_rounds)
        pattern_row.addWidget(self.round_search_btn)
        self.dday_final_btn = QPushButton("D-Day AI 최종조합")
        self.dday_final_btn.clicked.connect(self.show_dday_final_combinations)
        pattern_row.addWidget(self.dday_final_btn)
        lay.addLayout(pattern_row)

        self.rec_status = QLabel(
            "역대 Excel을 불러오면 자체추천이 자동 계산됩니다.\n"
            "추천조합·나온횟수·동반수·트리플·최근패턴은 입력번호 6개 이상이 필요합니다."
        )
        self.rec_status.setWordWrap(True)
        self.rec_status.setObjectName("card")
        lay.addWidget(self.rec_status)

        self.rec_table = QTableWidget(0, 13)
        self.rec_table.setHorizontalHeaderLabels(
            [
                "순위", "추천조합", "추천 이유", "신뢰도", "등급",
                "카테고리 점수", "종합 점수", "나온횟수", "동반수",
                "트리플", "최근패턴", "동반출현 횟수", "합계"
            ]
        )
        self.rec_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.rec_table.cellClicked.connect(self.show_recommend_detail)
        lay.addWidget(self.rec_table)

        self.detail_box = QPlainTextEdit()
        self.detail_box.setReadOnly(True)
        self.detail_box.setPlaceholderText(
            "추천조합을 클릭하면 동반출현·트리플·점수 상세가 표시됩니다."
        )
        self.detail_box.setMaximumHeight(170)
        lay.addWidget(self.detail_box)

        export = QPushButton("현재 순위 Excel 저장")
        export.clicked.connect(self.export_results)
        lay.addWidget(export)
        return p

    def show_dday_final_combinations(self) -> None:
        r = TKDDayFinalLab.RESULT
        combos = ["  ".join(f"{n:02d}" for n in c) for c in r["final5"]]
        lines = [
            "太炅 Lotto Lab D-Day AI 최종조합", "",
            f"검증 선택엔진: {r['protocol']['selected_preset']}",
            f"완전 보류검증: {r['protocol']['holdout']}",
            f"후보15 평균 포함: {r['holdout']['candidate15_avg']:.3f}개",
            f"후보15 5개 이상: {r['holdout']['candidate15_5plus']*100:.2f}%",
            f"후보15 6개 전부: {r['holdout']['candidate15_6all']}회",
            f"TOP5 최고조합 평균: {r['holdout']['top5_best_avg']:.3f}개",
            f"TOP5 3개 이상: {r['holdout']['top5_3plus']*100:.2f}%", "",
            "TOP 15 후보번호", " ".join(f"{n:02d}" for n in r["top15"]), "",
            "TOP 10", " ".join(f"{n:02d}" for n in r["top10"]), "",
            "TOP 7 핵심", " ".join(f"{n:02d}" for n in r["top7"]), "",
            "최종 추천조합 5개",
        ]
        lines += [f"{i+1}. {c}" for i,c in enumerate(combos)]
        lines += ["", r["disclaimer"]]
        QMessageBox.information(self, "D-Day AI 최종조합", "\n".join(lines))

    def _app_data_dir(self) -> Path:
        root = Path(os.getenv("LOCALAPPDATA", Path.home())) / "TaegyeongLottoLab"
        root.mkdir(parents=True, exist_ok=True)
        return root

    def _session_path(self) -> Path:
        return self._app_data_dir() / "last_session.json"

    def _merge_manual_draws(self) -> int:
        path = self._draw_data_path()
        if not path.exists() or not self.analyzer.draws:
            return 0
        try:
            records = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return 0

        existing = {draw.round_no for draw in self.analyzer.draws}
        added = 0
        for record in records:
            try:
                round_no = int(record["round_no"])
                numbers = tuple(sorted(int(n) for n in record["numbers"]))
                bonus = int(record["bonus"]) if record.get("bonus") is not None else None
                if round_no in existing:
                    continue
                if len(numbers) != 6 or len(set(numbers)) != 6:
                    continue
                if any(not 1 <= n <= 45 for n in numbers):
                    continue
                if bonus is not None and (not 1 <= bonus <= 45 or bonus in numbers):
                    continue
                self.analyzer.draws.append(Draw(round_no, numbers, bonus))
                existing.add(round_no)
                added += 1
            except Exception:
                continue

        if added:
            self.analyzer.draws.sort(key=lambda draw: draw.round_no)
            self.analyzer._analyze()
            self.pattern_cache.clear()
        return added

    def save_session(self) -> None:
        try:
            state = {
                "excel_path": getattr(self, "current_excel_path", ""),
                "source_input": self.source_input.toPlainText(),
                "fixed_input": self.fixed_input.text(),
                "excluded_input": self.excluded_input.text(),
                "candidate_input": self.candidate_input.text(),
                "photo_paths": [path for path in self.photo_paths if Path(path).exists()],
                "rec_category": self.rec_category.currentText(),
                "strategy": self.strategy_combo.currentText(),
                "mixed_preset": self.mixed_preset_combo.currentText(),
                "pattern_mode": self.pattern_mode_combo.currentText(),
                "round_search": self.round_search_input.text(),
                "stack_index": self.stack.currentIndex(),
                "window": {
                    "x": self.x(),
                    "y": self.y(),
                    "width": self.width(),
                    "height": self.height(),
                },
                "saved_at": datetime.now().isoformat(timespec="seconds"),
            }
            path = self._session_path()
            temp_path = path.with_suffix(".tmp")
            temp_path.write_text(
                json.dumps(state, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            temp_path.replace(path)
            self.statusBar().showMessage("현재 작업상태가 저장되었습니다.", 3000)
        except Exception as exc:
            self.statusBar().showMessage(f"상태 저장 실패: {exc}", 5000)

    @staticmethod
    def _set_combo_text(combo: QComboBox, text: str) -> None:
        index = combo.findText(text)
        if index >= 0:
            combo.setCurrentIndex(index)

    def restore_session(self) -> None:
        path = self._session_path()
        if not path.exists():
            self.statusBar().showMessage("역대 로또 Excel 파일을 불러오세요.")
            return

        try:
            state = json.loads(path.read_text(encoding="utf-8"))
        except Exception as exc:
            self.statusBar().showMessage(f"이전 상태를 읽지 못했습니다: {exc}")
            return

        self.suspend_auto_recommend = True
        widgets = [
            self.rec_category,
            self.strategy_combo,
            self.mixed_preset_combo,
            self.pattern_mode_combo,
        ]
        for widget in widgets:
            widget.blockSignals(True)

        try:
            self.source_input.setPlainText(state.get("source_input", ""))
            self.fixed_input.setText(state.get("fixed_input", ""))
            self.excluded_input.setText(state.get("excluded_input", ""))
            self.candidate_input.setText(state.get("candidate_input", ""))
            self.round_search_input.setText(state.get("round_search", ""))

            self._set_combo_text(
                self.rec_category, state.get("rec_category", "자체추천")
            )
            self._set_combo_text(
                self.strategy_combo, state.get("strategy", "균형형")
            )
            self._set_combo_text(
                self.mixed_preset_combo, state.get("mixed_preset", "균형형")
            )
            self._set_combo_text(
                self.pattern_mode_combo, state.get("pattern_mode", "자동종합")
            )

            self.photo_paths = [
                path for path in state.get("photo_paths", [])
                if Path(path).exists()
            ]
            self.photo_list.clear()
            for photo_path in self.photo_paths:
                self.photo_list.addItem(Path(photo_path).name)

            window = state.get("window", {})
            width = max(1100, int(window.get("width", 1320)))
            height = max(700, int(window.get("height", 850)))
            self.resize(width, height)
            if "x" in window and "y" in window:
                self.move(int(window["x"]), int(window["y"]))

            stack_index = int(state.get("stack_index", 0))
            if 0 <= stack_index < self.stack.count():
                self.stack.setCurrentIndex(stack_index)

            excel_path = state.get("excel_path", "")
            if excel_path and Path(excel_path).exists():
                self.current_excel_path = excel_path
                self.excel_progress.setValue(15)
                self.excel_status.setText(
                    f"이전 Excel 자동 복원 중: {Path(excel_path).name}"
                )
                QApplication.processEvents()
                self.analyzer.load_excel(excel_path)
                merged = self._merge_manual_draws()
                latest = self.analyzer.draws[-1].round_no
                self.excel_progress.setValue(100)
                self.excel_status.setText(
                    f"이전 Excel 자동 복원 완료: {Path(excel_path).name}\n"
                    f"분석 회차 {len(self.analyzer.draws):,}개 · 최신 {latest}회"
                    + (f" · 수동 추가 {merged}개 병합" if merged else "")
                )
                self.update_source_counts()
                self.statusBar().showMessage(
                    f"이전 작업상태 복원 완료 · 최신 {latest}회"
                )
            else:
                self.statusBar().showMessage(
                    "입력값과 설정은 복원했지만 이전 Excel 파일을 찾지 못했습니다."
                )
        except Exception as exc:
            self.statusBar().showMessage(f"이전 상태 복원 중 오류: {exc}", 7000)
        finally:
            for widget in widgets:
                widget.blockSignals(False)
            self.suspend_auto_recommend = False

        if self.analyzer.draws:
            self.generate_recommendations()

    def closeEvent(self, event) -> None:
        self.save_session()
        event.accept()

    def show_recommend_category(self, category: str) -> None:
        self.stack.setCurrentIndex(1)
        index = self.rec_category.findText(category)
        if index >= 0:
            self.rec_category.blockSignals(True)
            self.rec_category.setCurrentIndex(index)
            self.rec_category.blockSignals(False)
        self.generate_recommendations()

    def open_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "역대 로또 당첨번호 Excel 선택",
            "",
            "Excel (*.xlsx *.xls)",
        )
        if not path:
            return

        try:
            self.excel_progress.setValue(10)
            self.excel_status.setText(f"Excel 읽는 중: {Path(path).name}")
            self.statusBar().showMessage("역대 Excel을 읽고 있습니다...")
            QApplication.processEvents()

            self.analyzer.load_excel(path)
            self.current_excel_path = path
            merged_manual = self._merge_manual_draws()
            self.pattern_cache.clear()

            self.excel_progress.setValue(75)
            latest = self.analyzer.draws[-1].round_no
            QApplication.processEvents()

            self.excel_status.setText(
                f"Excel 등록 완료: {Path(path).name}\n"
                f"분석 회차 {len(self.analyzer.draws):,}개 · 최신 {latest}회 · "
                f"1등 조합 {len(self.analyzer.first_prize):,}개"
                + (f" · 수동 추가 {merged_manual}개 병합" if merged_manual else "")
            )
            self.excel_progress.setValue(100)

            # Excel만으로 계산 가능한 자체추천을 즉시 실행
            self.statusBar().showMessage("Excel 분석 완료 — 자체추천 계산 중...")
            self.show_recommend_category("자체추천")
            self.save_session()

            # 사진/직접입력 번호가 이미 있다면 다른 5개 항목도 사용할 수 있음을 표시
            try:
                counts = self.source_weights()
            except Exception:
                counts = Counter()

            if len(counts) >= 6:
                self.rec_status.setText(
                    "Excel과 입력번호가 모두 준비되었습니다. "
                    "왼쪽 항목을 누르면 각 기준 100조합이 계산됩니다."
                )

        except Exception as exc:
            self.excel_progress.setValue(0)
            self.excel_status.setText("Excel 등록 실패")
            QMessageBox.critical(
                self,
                "불러오기 오류",
                f"{exc}\n\n{traceback.format_exc(limit=3)}",
            )

    def prepare_ocr_image(self, image_path: str) -> tuple[str, str | None]:
        """고해상도 사진을 OCR에 충분한 크기로 축소해 처리시간을 줄입니다."""
        image = QImage(image_path)
        if image.isNull():
            return image_path, None

        max_side = max(image.width(), image.height())
        if max_side <= 1800:
            return image_path, None

        scaled = image.scaled(
            1800,
            1800,
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation,
        )
        temp_file = tempfile.NamedTemporaryFile(
            prefix="taegyeong_ocr_",
            suffix=".jpg",
            delete=False,
        )
        temp_path = temp_file.name
        temp_file.close()
        if not scaled.save(temp_path, "JPG", 88):
            try:
                os.unlink(temp_path)
            except OSError:
                pass
            return image_path, None
        return temp_path, temp_path

    def run_windows_ocr(self, image_path: str) -> list[int]:
        """Windows OCR을 축소 이미지와 캐시로 빠르게 호출합니다."""
        if sys.platform != "win32":
            raise RuntimeError("사진 OCR은 Windows 10/11에서만 사용할 수 있습니다.")

        source_path = Path(image_path).resolve()
        stat = source_path.stat()
        cache_key = (str(source_path), stat.st_mtime_ns, stat.st_size)
        if cache_key in self.ocr_cache:
            return list(self.ocr_cache[cache_key])

        prepared_path, temp_path = self.prepare_ocr_image(str(source_path))

        encoded = base64.b64encode(
            WINDOWS_OCR_PS.encode("utf-16le")
        ).decode("ascii")

        env = os.environ.copy()
        env["LOTTO_OCR_IMAGE"] = prepared_path

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
        result = [int(n) for n in numbers if 1 <= int(n) <= 45]
        self.ocr_cache[cache_key] = list(result)
        if temp_path:
            try:
                os.unlink(temp_path)
            except OSError:
                pass
        return result

    def append_ocr_numbers(self, numbers: list[int]) -> None:
        if not numbers:
            return
        current = self.source_input.toPlainText().rstrip()
        added = " ".join(map(str, numbers))
        self.source_input.blockSignals(True)
        self.source_input.setPlainText((current + "\n" + added).strip())
        self.source_input.blockSignals(False)
        self.update_source_counts()

    def set_ocr_controls_enabled(self, enabled: bool) -> None:
        self.photo_add_button.setEnabled(enabled)
        self.photo_delete_button.setEnabled(enabled)
        self.photo_rerun_button.setEnabled(enabled)

    def start_ocr(self, paths: list[str]) -> None:
        if self.ocr_worker is not None and self.ocr_worker.isRunning():
            QMessageBox.information(self, "OCR 진행 중", "현재 사진을 인식하고 있습니다. 완료 후 다시 시도하세요.")
            return

        self.suspend_auto_recommend = True
        self.set_ocr_controls_enabled(False)
        self.statusBar().showMessage(f"사진 OCR 준비 중 — 총 {len(paths)}장")

        self.ocr_worker = OCRWorker(paths, self.run_windows_ocr, self)
        self.ocr_worker.progress.connect(self.on_ocr_progress)
        self.ocr_worker.completed.connect(self.on_ocr_completed)
        self.ocr_worker.finished.connect(self.on_ocr_finished)
        self.ocr_worker.start()

    def on_ocr_progress(self, index: int, total: int, filename: str) -> None:
        self.statusBar().showMessage(f"사진 인식 중 {index}/{total}: {filename}")

    def on_ocr_completed(self, all_numbers: list, failures: list) -> None:
        if all_numbers:
            self.append_ocr_numbers([int(n) for n in all_numbers])
            self.statusBar().showMessage(
                f"사진 처리 완료 — 숫자 {len(all_numbers)}개 인식, 추천조합 계산 완료"
            )
        else:
            self.statusBar().showMessage("사진에서 1~45 숫자를 찾지 못했습니다.")

        message = f"인식된 숫자: {len(all_numbers)}개"
        if failures:
            message += "\n\n일부 오류:\n" + "\n".join(str(x) for x in failures[:5])
        QMessageBox.information(self, "사진 OCR 결과", message)

    def on_ocr_finished(self) -> None:
        self.suspend_auto_recommend = False
        self.set_ocr_controls_enabled(True)
        worker = self.ocr_worker
        self.ocr_worker = None
        if worker is not None:
            worker.deleteLater()

    def add_photos(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "번호 사진 선택", "",
            "Images (*.png *.jpg *.jpeg *.webp *.bmp)"
        )
        if not paths:
            return

        for path in paths:
            if path not in self.photo_paths:
                self.photo_paths.append(path)
                self.photo_list.addItem(Path(path).name)

        self.start_ocr(paths)

    def rerun_selected_photo_ocr(self) -> None:
        row = self.photo_list.currentRow()
        if row < 0:
            QMessageBox.information(self, "사진 선택", "다시 인식할 사진을 선택하세요.")
            return
        self.start_ocr([self.photo_paths[row]])

    def delete_photo(self) -> None:
        row = self.photo_list.currentRow()
        if row >= 0:
            self.photo_list.takeItem(row)
            self.photo_paths.pop(row)

    @staticmethod
    def _parse_special_numbers(text: str, label: str, maximum: int = 10) -> tuple[int, ...]:
        numbers = sorted({
            int(x)
            for x in re.findall(r"\d{1,2}", text)
            if 1 <= int(x) <= 45
        })
        if len(numbers) > maximum:
            raise ValueError(f"{label}는 최대 {maximum}개까지 입력할 수 있습니다.")
        return tuple(numbers)

    def fixed_numbers(self) -> tuple[int, ...]:
        text = self.fixed_input.text() if hasattr(self, "fixed_input") else ""
        return self._parse_special_numbers(text, "필수번호", 5)

    def excluded_numbers(self) -> tuple[int, ...]:
        text = self.excluded_input.text() if hasattr(self, "excluded_input") else ""
        return self._parse_special_numbers(text, "제외번호", 10)

    def candidate_numbers(self) -> tuple[int, ...]:
        text = self.candidate_input.text() if hasattr(self, "candidate_input") else ""
        return self._parse_special_numbers(text, "후보번호", 10)

    def validate_special_numbers(
        self,
        source_weights: Counter[int],
        fixed_numbers: tuple[int, ...],
        excluded_numbers: tuple[int, ...],
        candidate_numbers: tuple[int, ...],
    ) -> None:
        fixed_set = set(fixed_numbers)
        excluded_set = set(excluded_numbers)
        candidate_set = set(candidate_numbers)

        if fixed_set & excluded_set:
            raise ValueError(
                "필수번호와 제외번호에 같은 번호가 있습니다: "
                + ", ".join(map(str, sorted(fixed_set & excluded_set)))
            )
        if fixed_set & candidate_set:
            raise ValueError(
                "필수번호와 후보번호에 같은 번호가 있습니다: "
                + ", ".join(map(str, sorted(fixed_set & candidate_set)))
            )
        if excluded_set & candidate_set:
            raise ValueError(
                "제외번호와 후보번호에 같은 번호가 있습니다: "
                + ", ".join(map(str, sorted(excluded_set & candidate_set)))
            )

        missing_fixed = [n for n in fixed_numbers if n not in source_weights]
        if missing_fixed:
            raise ValueError(
                "필수번호는 일반 번호 입력란에도 포함되어야 합니다: "
                + ", ".join(map(str, missing_fixed))
            )

    def source_weights(self) -> Counter[int]:
        # 필수·제외·후보번호 입력란은 여기 합산하지 않습니다.
        # 따라서 일반 입력번호의 나온횟수가 중복 증가하지 않습니다.
        nums = parse_numbers(self.source_input.toPlainText())
        return Counter(nums)

    def update_source_counts(self) -> None:
        try:
            counts = self.source_weights()
            if not counts:
                self.source_summary.setText("입력된 번호가 없습니다.")
                return
            ranked = sorted(counts.items(), key=lambda x: (-x[1], x[0]))
            fixed = self.fixed_numbers()
            excluded = self.excluded_numbers()
            candidate = self.candidate_numbers()

            fixed_text = "없음" if not fixed else ", ".join(map(str, fixed))
            excluded_text = "없음" if not excluded else ", ".join(map(str, excluded))
            candidate_text = "없음" if not candidate else ", ".join(map(str, candidate))

            self.source_summary.setText(
                f"고유 번호 {len(counts)}개 / 전체 입력 {sum(counts.values())}개\n"
                f"필수번호: {fixed_text} / 제외번호: {excluded_text} / 후보번호: {candidate_text}\n"
                + " · ".join(f"{n}번 {c}회" for n, c in ranked)
            )
            if self.analyzer.draws and len(counts) >= 6:
                self.excel_status.setText(
                    self.excel_status.text()
                    + "\n입력번호 준비 완료 — 모든 추천 항목을 사용할 수 있습니다."
                )
                # 여러 사진 처리 중에는 매 사진마다 100조합을 재계산하지 않습니다.
                if not self.suspend_auto_recommend:
                    self.show_recommend_category("추천조합")
            elif not self.analyzer.draws:
                self.source_summary.setText(
                    self.source_summary.text()
                    + "\n역대 Excel을 불러오면 추천 계산이 시작됩니다."
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

    def generate_recommendations(self, *_args) -> None:
        if not self.analyzer.draws:
            if hasattr(self, "rec_table"):
                self.rec_table.setRowCount(0)
            if hasattr(self, "rec_status"):
                self.rec_status.setText(
                    "먼저 왼쪽 아래의 '역대 Excel 불러오기'로 당첨번호 파일을 등록하세요."
                )
            self.statusBar().showMessage("역대 Excel이 필요합니다.")
            return
        try:
            recommender = Recommender(self.analyzer)
            category = self.rec_category.currentText()

            if category == "자체추천":
                mode = self.pattern_mode_combo.currentText()
                self.statusBar().showMessage(
                    "v27 후보생성 → v40 조합엔진으로 자체추천 100조합 계산 중..."
                )
                self.rec_status.setText(
                    "자체추천 계산 중 · 후보: v27 페어다양성100 · 조합: v40 Elite Survival"
                )
                QApplication.processEvents()
                self.recommendations = recommender.generate_self_v27_v40(100, mode=mode)
                fixed_numbers = ()
                excluded_numbers = ()
                candidate_numbers = ()
                strategy = "자체추천-v27후보-v40조합"
            elif category == "AI이중후보추천":
                fixed_numbers = self.fixed_numbers()
                excluded_numbers = self.excluded_numbers()
                candidate_numbers = self.candidate_numbers()
                self.rec_status.setText(
                    "AI이중후보추천 계산 중 · 기존 핵심20 + 공격형 보조3 = 후보23개"
                )
                QApplication.processEvents()
                self.recommendations = TKDualCandidateLab.generate(
                    self.analyzer,
                    1000,
                    fixed_numbers=fixed_numbers,
                    excluded_numbers=excluded_numbers,
                    candidate_numbers=candidate_numbers,
                )
                strategy = "AI이중후보-공격형23"
            elif category == "AI후보번호연구":
                fixed_numbers = self.fixed_numbers()
                excluded_numbers = self.excluded_numbers()
                candidate_numbers = self.candidate_numbers()
                self.rec_status.setText(
                    "AI후보번호연구 계산 중 · 검증된 후보 15개를 먼저 선정한 뒤 100조합을 생성합니다."
                )
                QApplication.processEvents()
                self.recommendations = TKCandidateLab.generate(
                    self.analyzer,
                    1000,
                    fixed_numbers=fixed_numbers,
                    excluded_numbers=excluded_numbers,
                    candidate_numbers=candidate_numbers,
                )
                strategy = "AI후보번호연구-15개"
            elif category == "AI연구소추천":
                fixed_numbers = self.fixed_numbers()
                excluded_numbers = self.excluded_numbers()
                candidate_numbers = self.candidate_numbers()
                self.rec_status.setText("AI연구소추천 계산 중 · 엔진 월드컵·메타앙상블 결과 적용")
                QApplication.processEvents()
                self.recommendations = TKV8EvolutionLab.generate(
                    self.analyzer, 1000, fixed_numbers=fixed_numbers,
                    excluded_numbers=excluded_numbers, candidate_numbers=candidate_numbers,
                )
                strategy = "V8-AI연구소"
            elif category == "AI진화추천":
                fixed_numbers = self.fixed_numbers()
                excluded_numbers = self.excluded_numbers()
                candidate_numbers = self.candidate_numbers()
                self.rec_status.setText(
                    "AI진화추천 계산 중 · 1~1000회 초기학습 후 1001회부터 최신까지 순차 성장 가중치 적용"
                )
                self.statusBar().showMessage("AI진화추천 100조합 계산 중...")
                QApplication.processEvents()
                self.recommendations = TKEvolutionEngine.generate(
                    self.analyzer,
                    1000,
                    fixed_numbers=fixed_numbers,
                    excluded_numbers=excluded_numbers,
                    candidate_numbers=candidate_numbers,
                )
                strategy = "AI회차별성장"
            elif category == "성과최적추천":
                fixed_numbers = self.fixed_numbers()
                excluded_numbers = self.excluded_numbers()
                candidate_numbers = self.candidate_numbers()
                self.rec_status.setText(
                    "성과최적추천 계산 중 · 30,000개 설정 경쟁에서 선택된 가중치 적용"
                )
                self.statusBar().showMessage("성과최적추천 100조합 계산 중...")
                QApplication.processEvents()
                self.recommendations = TKPerformanceEngine.generate(
                    self.analyzer,
                    1000,
                    fixed_numbers=fixed_numbers,
                    excluded_numbers=excluded_numbers,
                    candidate_numbers=candidate_numbers,
                )
                strategy = "성과최적엔진"
            elif category == "특이패턴추천":
                fixed_numbers = self.fixed_numbers()
                excluded_numbers = self.excluded_numbers()
                candidate_numbers = self.candidate_numbers()
                source_weights = self.source_weights()
                self.validate_special_numbers(
                    source_weights,
                    fixed_numbers,
                    excluded_numbers,
                    candidate_numbers,
                )
                mode = self.pattern_mode_combo.currentText()
                self.rec_status.setText(
                    f"특이패턴추천 계산 중 · {mode} · 패턴투표와 검증점수를 종합합니다."
                )
                self.statusBar().showMessage("특이패턴 추천 100조합 계산 중...")
                QApplication.processEvents()
                self.recommendations = recommender.generate_pattern(
                    1000,
                    mode=mode,
                    fixed_numbers=fixed_numbers,
                    excluded_numbers=excluded_numbers,
                    candidate_numbers=candidate_numbers,
                )
                strategy = f"특이패턴-{mode}"
            elif category == "통합데이터추천":
                weights = self.source_weights()
                fixed_numbers = self.fixed_numbers()
                excluded_numbers = self.excluded_numbers()
                candidate_numbers = self.candidate_numbers()
                self.validate_special_numbers(
                    weights,
                    fixed_numbers,
                    excluded_numbers,
                    candidate_numbers,
                )
                if len(weights) < 6:
                    self.rec_table.setRowCount(0)
                    self.rec_status.setText(
                        "통합데이터추천은 입력번호가 최소 6개 필요합니다."
                    )
                    return
                preset = self.mixed_preset_combo.currentText()
                self.rec_status.setText(
                    f"통합데이터추천 계산 중 · 입력번호 + 최근100/500/1000회 · {preset}"
                )
                self.statusBar().showMessage("통합데이터추천 100조합 계산 중...")
                QApplication.processEvents()
                self.recommendations = recommender.generate_mixed(
                    weights,
                    1000,
                    preset,
                    fixed_numbers=fixed_numbers,
                    excluded_numbers=excluded_numbers,
                    candidate_numbers=candidate_numbers,
                )
                strategy = f"통합-{preset}"
            else:
                weights = self.source_weights()
                fixed_numbers = self.fixed_numbers()
                excluded_numbers = self.excluded_numbers()
                candidate_numbers = self.candidate_numbers()
                strategy = self.strategy_combo.currentText()
                self.validate_special_numbers(
                    weights,
                    fixed_numbers,
                    excluded_numbers,
                    candidate_numbers,
                )
                if len(weights) < 6:
                    self.rec_table.setRowCount(0)
                    self.rec_status.setText(
                        f"{category}은 사진 또는 직접 입력에서 고유 번호 6개 이상이 필요합니다."
                    )
                    self.statusBar().showMessage(
                        f"{category}: 사진 또는 직접 입력으로 고유 번호 6개 이상을 입력하세요."
                    )
                    return
                filter_name, filter_desc = recommender.filter_mode(len(weights))
                self.rec_status.setText(
                    f"{category} 100조합 계산 중 · {filter_desc} · 전략: {strategy}"
                )
                self.statusBar().showMessage(
                    f"{category} 계산 중 · {filter_name} 필터"
                )
                QApplication.processEvents()
                self.recommendations = recommender.generate(
                    weights,
                    1000,
                    20,
                    300,
                    True,
                    category,
                    fixed_numbers=fixed_numbers,
                    excluded_numbers=excluded_numbers,
                    candidate_numbers=candidate_numbers,
                    strategy=strategy,
                )
            # 자체추천은 내부에서 v27 후보생성 → v40 최종선별을 완료합니다.
            # 그 외 모든 항목은 넓은 후보조합에서 v40 Elite Survival이 최종 100개를 선별합니다.
            if category != "자체추천":
                self.recommendations = Recommender.select_diverse(
                    list(self.recommendations), 100
                )
                strategy = f"{strategy}-v40최종조합"

            if not self.recommendations:
                QMessageBox.information(
                    self, "결과 없음",
                    "추천 가능한 조합이 없습니다. 입력번호를 10개 이상으로 확인해 주세요."
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
                reason = recommender.recommendation_reason(
                    metrics,
                    fixed_numbers,
                    candidate_numbers,
                    combo,
                )
                if metrics.get("pattern_names"):
                    reason = (
                        "패턴투표: " + ", ".join(metrics["pattern_names"])
                        + " / " + reason
                    )
                if metrics.get("performance_reasons"):
                    reason = (
                        "성과최적: " + " / ".join(metrics["performance_reasons"][:2])
                    )
                confidence = recommender.confidence_score(score, metrics)
                grade = recommender.confidence_grade(confidence)
                rank_text = f"TOP {r}" if r <= 10 else str(r)
                values = [
                    rank_text,
                    " · ".join(map(str, combo)),
                    reason,
                    f"{confidence:.1f}",
                    grade,
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
                    "composite": 6, "input": 7, "pair": 8,
                    "triple": 9, "recent": 10, "mixed": 5,
                    "pattern": 5, "performance": 5, "evolution": 5, "evolution_lab": 5, "candidate_lab": 5, "dual_candidate": 5, "self": 5,
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

                if r <= 10:
                    for top_col in range(self.rec_table.columnCount()):
                        top_item = self.rec_table.item(r - 1, top_col)
                        if top_item is not None:
                            top_item.setBackground(QBrush(QColor("#3A3215")))

                combo_item = self.rec_table.item(r - 1, 1)
                if metric_value >= 70:
                    combo_item.setForeground(QBrush(QColor("#7CFF8A")))
                elif metric_value >= 50:
                    combo_item.setForeground(QBrush(QColor("#FFD95A")))

            self.rec_table.resizeColumnsToContents()
            if category == "자체추천":
                first_metrics = self.recommendations[0][2]
                filter_text = (
                    f"현재패턴 {first_metrics.get('self_pattern_mode', '')} · "
                    f"패턴비중 {first_metrics.get('self_pattern_mix', 0) * 100:.0f}% · "
                    f"패턴TOP10 {' · '.join(map(str, first_metrics.get('self_pattern_top10', [])))} · "
                    f"후보풀 {' · '.join(map(str, first_metrics.get('self_candidate_pool', [])))}"
                )
            elif category == "통합데이터추천":
                first_metrics = self.recommendations[0][2]
                filter_text = (
                    f"프리셋 {first_metrics.get('mixed_preset', '')} · "
                    f"입력 {first_metrics.get('mixed_input_weight', 0) * 100:.0f}% / "
                    f"최근100 {first_metrics.get('mixed_100_weight', 0) * 100:.0f}% / "
                    f"최근500 {first_metrics.get('mixed_500_weight', 0) * 100:.0f}% / "
                    f"최근1000 {first_metrics.get('mixed_1000_weight', 0) * 100:.0f}% · "
                    f"후보풀 {' · '.join(map(str, first_metrics.get('mixed_candidate_pool', [])))}"
                )
            else:
                mode_values = {
                    str(row_metrics.get("filter_mode", "자동"))
                    for _, _, row_metrics in self.recommendations
                }
                filter_text = ", ".join(sorted(mode_values))

            self.rec_status.setText(
                f"{category} 기준 {len(self.recommendations)}조합 계산 완료 · "
                f"{filter_text} · 조합엔진 v40 Elite Survival 적용"
            )
            self.statusBar().showMessage(
                f"{category} 추천 {len(self.recommendations)}개 완료"
            )
        except Exception as e:
            QMessageBox.warning(self, "추천 오류", str(e))



    def on_mixed_preset_changed(self, preset: str) -> None:
        self.recommendations = []
        if hasattr(self, "rec_table"):
            self.rec_table.setRowCount(0)
        self.rec_status.setText(
            f"통합데이터 설정 변경: {preset} · 후보번호와 추천조합을 다시 계산합니다."
        )
        self.statusBar().showMessage(
            f"{preset} 적용 중 · 이전 통합추천 결과를 재사용하지 않습니다."
        )
        if self.suspend_auto_recommend or not self.analyzer.draws:
            return
        if self.rec_category.currentText() == "통합데이터추천":
            QApplication.processEvents()
            self.generate_recommendations()
        else:
            self.save_session()

    def on_pattern_mode_changed(self, mode: str) -> None:
        """패턴 변경 시 기존 결과를 폐기하고 새 조합을 강제로 계산합니다."""
        self.pattern_cache.clear()
        self.recommendations = []
        if hasattr(self, "rec_table"):
            self.rec_table.setRowCount(0)
        self.rec_status.setText(
            f"패턴 변경: {mode} · 후보번호와 추천조합을 새로 계산합니다."
        )
        self.statusBar().showMessage(
            f"{mode} 적용 중 · 이전 추천결과를 재사용하지 않습니다."
        )
        if self.suspend_auto_recommend or not self.analyzer.draws:
            return
        if self.rec_category.currentText() in ("자체추천", "특이패턴추천"):
            QApplication.processEvents()
            self.generate_recommendations()
        else:
            self.save_session()

    def show_boundary_companion_review_report(self) -> None:
        r = TKBoundaryCompanionReviewLab.RESULT
        base = r["baseline_attack25"]
        new = r["boundary_review"]
        lines = [
            "太炅 Boundary Companion Review v10.6", "",
            f"완전 보류: {r['protocol']['holdout']}",
            f"경계폭: {r['protocol']['selected_boundary_width']}개",
            f"페어/트리플 가중치: {r['protocol']['selected_pair_weight']:.2f} / "
            f"{r['protocol']['selected_triple_weight']:.2f}",
            f"최종 후보수: {r['protocol']['selected_candidate_size']}개", "",
            "[기존 공격형25]",
            f"평균 포함: {base['average_hits']:.3f}",
            f"4개 이상: {base['four_plus_rate']*100:.2f}%",
            f"5개 이상: {base['five_plus_rate']*100:.2f}%",
            f"6개 전부: {base['six_all_cases']}회", "",
            "[경계번호 재심사]",
            f"평균 포함: {new['average_hits']:.3f}",
            f"4개 이상: {new['four_plus_rate']*100:.2f}%",
            f"5개 이상: {new['five_plus_rate']*100:.2f}%",
            f"6개 전부: {new['six_all_cases']}회",
            f"변경 회차: {new['changed_rounds']}회",
            f"경계 교체 순증 적중: {new['net_boundary_hits']:+d}개", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
        ]
        QMessageBox.information(self, "V10.6 경계번호 동반수 검증", "\n".join(lines))

    def show_survival_competition_report(self) -> None:
        r = TKSurvivalCompetitionLab.RESULT
        base = r["baseline_attack25"]
        new = r["survival_competition"]
        lines = [
            "太炅 Survival Competition Lab v10.5", "",
            f"완전 보류: {r['protocol']['holdout']}",
            f"선택 후보수: {r['protocol']['selected_candidate_size']}개",
            f"강제 제외수: {r['protocol']['selected_hard_exclude']}개",
            f"제외 가중치: {r['protocol']['selected_exclude_weight']:.2f}", "",
            "[기존 공격형25]",
            f"평균 포함: {base['average_hits']:.3f}",
            f"4개 이상: {base['four_plus_rate']*100:.2f}%",
            f"5개 이상: {base['five_plus_rate']*100:.2f}%",
            f"6개 전부: {base['six_all_cases']}회", "",
            "[후보 생존경쟁]",
            f"평균 후보수: {new['average_candidate_size']:.2f}",
            f"평균 포함: {new['average_hits']:.3f}",
            f"4개 이상: {new['four_plus_rate']*100:.2f}%",
            f"5개 이상: {new['five_plus_rate']*100:.2f}%",
            f"6개 전부: {new['six_all_cases']}회", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
        ]
        QMessageBox.information(self, "V10.5 후보생존·제외AI 검증", "\n".join(lines))

    def show_core23_aux_layer_report(self) -> None:
        r = TKCore23AuxLayerLab.RESULT
        core = r["stable_core23"]
        new = r["core23_aux_layer"]
        aggr = r["aggressive25_reference"]
        lines = [
            "太炅 Core23 + Aux Layer v10.4", "",
            f"완전 보류: {r['protocol']['holdout']}",
            f"핵심후보: {r['protocol']['core_size']}개",
            f"보조후보 최대: {r['protocol']['selected_aux_count']}개",
            f"활성 점수차: {r['protocol']['selected_min_gap']:.2f}", "",
            "[핵심23]",
            f"평균 포함: {core['average_hits']:.3f}",
            f"4개 이상: {core['four_plus_rate']*100:.2f}%",
            f"5개 이상: {core['five_plus_rate']*100:.2f}%",
            f"6개 전부: {core['six_all_cases']}회", "",
            "[공격형25 참고]",
            f"평균 포함: {aggr['average_hits']:.3f}",
            f"5개 이상: {aggr['five_plus_rate']*100:.2f}%",
            f"6개 전부: {aggr['six_all_cases']}회", "",
            "[핵심23+보조층]",
            f"평균 총후보: {new['average_total_size']:.2f}",
            f"평균 포함: {new['average_hits']:.3f}",
            f"4개 이상: {new['four_plus_rate']*100:.2f}%",
            f"5개 이상: {new['five_plus_rate']*100:.2f}%",
            f"6개 전부: {new['six_all_cases']}회",
            f"보조후보 적중회차: {new['aux_hit_rate']*100:.2f}%", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
        ]
        QMessageBox.information(self, "V10.4 핵심23+보조2 검증", "\n".join(lines))

    def show_stable_aggressive_meta_gate_report(self) -> None:
        r = TKStableAggressiveMetaGateLab.RESULT
        stable = r["stable_v10_1"]
        aggr = r["aggressive_v10_2"]
        meta = r["meta_gate_v10_3"]
        lines = [
            "太炅 Stable/Aggressive Meta Gate v10.3", "",
            f"완전 보류: {r['protocol']['holdout']}",
            f"선택 모델: {r['protocol']['selected_gate_model']}",
            f"선택 임계값: {r['protocol']['selected_threshold']:.2f}", "",
            "[안정형 v10.1]",
            f"평균 포함: {stable['average_hits']:.3f}",
            f"4개 이상: {stable['four_plus_rate']*100:.2f}%",
            f"5개 이상: {stable['five_plus_rate']*100:.2f}%",
            f"6개 전부: {stable['six_all_cases']}회", "",
            "[공격형 v10.2]",
            f"평균 포함: {aggr['average_hits']:.3f}",
            f"4개 이상: {aggr['four_plus_rate']*100:.2f}%",
            f"5개 이상: {aggr['five_plus_rate']*100:.2f}%",
            f"6개 전부: {aggr['six_all_cases']}회", "",
            "[메타게이트 v10.3]",
            f"평균 포함: {meta['average_hits']:.3f}",
            f"4개 이상: {meta['four_plus_rate']*100:.2f}%",
            f"5개 이상: {meta['five_plus_rate']*100:.2f}%",
            f"6개 전부: {meta['six_all_cases']}회",
            f"선택 횟수: {meta['choices']}", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
        ]
        QMessageBox.information(self, "V10.3 메타게이트 검증", "\n".join(lines))

    def show_ensemble_exclusion_ai_report(self) -> None:
        r = TKEnsembleExclusionAILab.RESULT
        base = r["baseline_v10_1"]
        new = r["ensemble_exclusion_ai"]
        lines = [
            "太炅 Ensemble Exclusion AI Lab v10.2", "",
            f"검증: {r['protocol']['holdout']}",
            f"비선형 가중치: {r['protocol']['selected_tree_weight']:.2f}",
            f"제외 패널티: {r['protocol']['selected_exclude_penalty']:.2f}", "",
            "[v10.1 기준]",
            f"평균 포함: {base['average_hits']:.3f}",
            f"4개 이상: {base['four_plus_rate']*100:.2f}%",
            f"5개 이상: {base['five_plus_rate']*100:.2f}%",
            f"6개 전부: {base['six_all_cases']}회", "",
            "[v10.2 앙상블 제외AI]",
            f"평균 포함: {new['average_hits']:.3f}",
            f"4개 이상: {new['four_plus_rate']*100:.2f}%",
            f"5개 이상: {new['five_plus_rate']*100:.2f}%",
            f"6개 전부: {new['six_all_cases']}회", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
        ]
        QMessageBox.information(self, "V10.2 앙상블 제외AI 검증", "\n".join(lines))

    def show_regime_candidate_size_report(self) -> None:
        r = TKRegimeCandidateSizeLab.RESULT
        base = r["baseline_fixed23"]
        new = r["regime_dynamic_size"]
        lines = [
            "太炅 Regime Candidate Size Lab v10.1", "",
            f"검증: {r['protocol']['holdout']}",
            f"유형별 선택 후보수: {r['protocol']['selected_sizes']}", "",
            "[고정 후보23]",
            f"평균 후보수: {base['average_candidate_size']:.2f}",
            f"평균 포함: {base['average_hits']:.3f}",
            f"5개 이상: {base['five_plus_rate']*100:.2f}%",
            f"6개 전부: {base['six_all_cases']}회", "",
            "[유형별 후보수]",
            f"평균 후보수: {new['average_candidate_size']:.2f}",
            f"평균 포함: {new['average_hits']:.3f}",
            f"5개 이상: {new['five_plus_rate']*100:.2f}%",
            f"6개 전부: {new['six_all_cases']}회", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
        ]
        QMessageBox.information(self, "V10.1 유형별 후보수 검증", "\n".join(lines))

    def show_regime_candidate_report(self) -> None:
        r = TKRegimeCandidateLab.RESULT
        base = r["baseline_global_candidate23"]
        new = r["regime_candidate23"]
        lines = [
            "太炅 Regime Candidate Lab v10.0", "",
            f"학습: {r['protocol']['training']}",
            f"조정검증: {r['protocol']['validation']}",
            f"완전 보류: {r['protocol']['holdout']}",
            f"유형 수: {r['protocol']['regimes']}개",
            f"선택 유형가중치: {r['protocol']['selected_regime_weight']:.2f}", "",
            "[글로벌 후보23]",
            f"평균 포함: {base['average_hits']:.3f}개",
            f"4개 이상: {base['four_plus_rate']*100:.1f}%",
            f"5개 이상: {base['five_plus_rate']*100:.1f}%",
            f"6개 전부: {base['six_all_cases']}회", "",
            "[회차유형 전용 후보23]",
            f"평균 포함: {new['average_hits']:.3f}개",
            f"4개 이상: {new['four_plus_rate']*100:.1f}%",
            f"5개 이상: {new['five_plus_rate']*100:.1f}%",
            f"6개 전부: {new['six_all_cases']}회", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
        ]
        QMessageBox.information(self, "V10 회차유형 후보 검증", "\n".join(lines))

    def show_safe_recovery_layer_report(self) -> None:
        r = TKSafeRecoveryLayerLab.RESULT
        base = r["baseline_70_30"]
        new = r["safe_recovery_70_30"]
        lines = [
            "太炅 Safe Recovery Layer Lab v9.6", "",
            f"조정검증: {r['protocol']['validation']}",
            f"완전 보류: {r['protocol']['holdout']}",
            f"기본후보: {r['protocol']['base_candidate_size']}개",
            f"안정형/공격형: {r['protocol']['stable_combinations']} · {r['protocol']['aggressive_combinations']}",
            f"선택 보조후보 수: {r['protocol']['selected_aux_count']}개", "",
            "[기존 70:30]",
            f"최고조합 평균: {base['best_hit_average']:.3f}개",
            f"4개 이상: {base['four_plus_rate']*100:.1f}%",
            f"5개 이상: {base['five_plus_rate']*100:.1f}%",
            f"6개 전부: {base['six_all_cases']}회", "",
            "[안전 복구층]",
            f"최고조합 평균: {new['best_hit_average']:.3f}개",
            f"4개 이상: {new['four_plus_rate']*100:.1f}%",
            f"5개 이상: {new['five_plus_rate']*100:.1f}%",
            f"6개 전부: {new['six_all_cases']}회", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
        ]
        QMessageBox.information(self, "안전 복구층 검증", "\n".join(lines))

    def show_multi_failure_experts_report(self) -> None:
        r = TKMultiFailureExpertsLab.RESULT
        base = r["baseline_candidate23"]
        new = r["multi_failure_candidate23"]
        lines = [
            "太炅 Multi Failure Experts Lab v9.5", "",
            f"학습: {r['protocol']['training']}",
            f"조정검증: {r['protocol']['validation']}",
            f"완전 보류: {r['protocol']['holdout']}",
            f"전문가: {' · '.join(r['protocol']['experts'])}",
            f"최소 합의/임계값: {r['protocol']['selected_min_votes']}명 · {r['protocol']['selected_threshold']:.3f}", "",
            "[기존 후보23]",
            f"평균 포함: {base['average_hits']:.3f}개",
            f"4개 이상: {base['four_plus_rate']*100:.1f}%",
            f"5개 이상: {base['five_plus_rate']*100:.1f}%",
            f"6개 전부: {base['six_all_cases']}회", "",
            "[다중 실패전문가 후보23]",
            f"교체 적용 회차: {new['changed_rounds']}회",
            f"평균 포함: {new['average_hits']:.3f}개",
            f"4개 이상: {new['four_plus_rate']*100:.1f}%",
            f"5개 이상: {new['five_plus_rate']*100:.1f}%",
            f"6개 전부: {new['six_all_cases']}회", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
        ]
        QMessageBox.information(self, "다중 실패전문가 검증", "\n".join(lines))

    def show_failure_recovery_report(self) -> None:
        r = TKFailureRecoveryLab.RESULT
        base = r["baseline_candidate23"]
        new = r["failure_recovery_candidate23"]
        lines = [
            "太炅 Failure Recovery Lab v9.4", "",
            f"학습: {r['protocol']['training']}",
            f"조정검증: {r['protocol']['validation']}",
            f"완전 보류: {r['protocol']['holdout']}",
            f"피처: {r['protocol']['features']}개",
            f"선택 교체수/임계값: {r['protocol']['selected_swap_count']} · "
            f"{r['protocol']['selected_threshold']:.3f}", "",
            "[기존 후보23]",
            f"평균 포함: {base['average_hits']:.3f}개",
            f"4개 이상: {base['four_plus_rate']*100:.1f}%",
            f"5개 이상: {base['five_plus_rate']*100:.1f}%",
            f"6개 전부: {base['six_all_cases']}회", "",
            "[실패회차 복구 후보23]",
            f"총 교체: {new['total_swaps']}개",
            f"평균 포함: {new['average_hits']:.3f}개",
            f"4개 이상: {new['four_plus_rate']*100:.1f}%",
            f"5개 이상: {new['five_plus_rate']*100:.1f}%",
            f"6개 전부: {new['six_all_cases']}회", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
        ]
        QMessageBox.information(self, "실패회차 복구 검증", "\n".join(lines))

    def show_dynamic_portfolio_report(self) -> None:
        r = TKDynamicPortfolioLab.RESULT
        base = r["baseline_fixed_70_30"]
        new = r["dynamic_portfolio"]
        lines = [
            "太炅 Dynamic Portfolio Lab v9.3", "",
            f"학습: {r['protocol']['training']}",
            f"조정검증: {r['protocol']['validation']}",
            f"완전 보류: {r['protocol']['holdout']}",
            f"회차유형 피처: {r['protocol']['round_type_features']}개",
            f"선택 모델: {r['protocol']['selected_model']}", "",
            "[고정 70:30]",
            f"최고조합 평균: {base['best_hit_average']:.3f}개",
            f"4개 이상: {base['four_plus_rate']*100:.1f}%",
            f"5개 이상: {base['five_plus_rate']*100:.1f}%",
            f"6개 전부: {base['six_all_cases']}회", "",
            "[회차유형 동적 포트폴리오]",
            f"최고조합 평균: {new['best_hit_average']:.3f}개",
            f"4개 이상: {new['four_plus_rate']*100:.1f}%",
            f"5개 이상: {new['five_plus_rate']*100:.1f}%",
            f"6개 전부: {new['six_all_cases']}회",
            f"선택 횟수: {new['portfolio_counts']}", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
        ]
        QMessageBox.information(self, "회차유형 동적포트폴리오 검증", "\n".join(lines))

    def show_dual_portfolio_report(self) -> None:
        r = TKDualPortfolioLab.RESULT
        base = r["baseline_top100"]
        new = r["dual_portfolio_100"]
        lines = [
            "太炅 Dual Portfolio Lab v9.2", "",
            f"조정검증: {r['protocol']['validation']}",
            f"완전 보류: {r['protocol']['holdout']}",
            f"안정형/공격형: {r['protocol']['selected_stable_count']} · "
            f"{r['protocol']['selected_aggressive_count']}",
            f"중복 제한: {r['protocol']['selected_overlap_limit']}개", "",
            "[기존 상위100조합]",
            f"최고조합 평균: {base['best_hit_average']:.3f}개",
            f"4개 이상: {base['four_plus_rate']*100:.1f}%",
            f"5개 이상: {base['five_plus_rate']*100:.1f}%",
            f"6개 전부: {base['six_all_cases']}회", "",
            "[이중 포트폴리오100조합]",
            f"최고조합 평균: {new['best_hit_average']:.3f}개",
            f"4개 이상: {new['four_plus_rate']*100:.1f}%",
            f"5개 이상: {new['five_plus_rate']*100:.1f}%",
            f"6개 전부: {new['six_all_cases']}회", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
        ]
        QMessageBox.information(self, "이중 포트폴리오 검증", "\n".join(lines))

    def show_combination_exposure_report(self) -> None:
        r = TKCombinationExposureLab.RESULT
        base = r["baseline_top100"]
        new = r["exposure_optimized_100"]
        lines = [
            "太炅 Combination Exposure Lab v9.1", "",
            f"조정검증: {r['protocol']['validation']}",
            f"완전 보류: {r['protocol']['holdout']}",
            f"후보: {r['protocol']['candidate_size']}개 · 조합: {r['protocol']['combinations_per_round']}개",
            f"선택 alpha/노출/중복: {r['protocol']['selected_alpha']} · "
            f"{r['protocol']['selected_exposure_weight']} · {r['protocol']['selected_overlap_weight']}", "",
            "[기존 상위100조합]",
            f"최고조합 평균: {base['best_hit_average']:.3f}개",
            f"4개 이상: {base['four_plus_rate']*100:.1f}%",
            f"5개 이상: {base['five_plus_rate']*100:.1f}%",
            f"6개 전부: {base['six_all_cases']}회", "",
            "[노출분배 최적화100조합]",
            f"최고조합 평균: {new['best_hit_average']:.3f}개",
            f"4개 이상: {new['four_plus_rate']*100:.1f}%",
            f"5개 이상: {new['five_plus_rate']*100:.1f}%",
            f"6개 전부: {new['six_all_cases']}회", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
        ]
        QMessageBox.information(self, "조합 노출분배 검증", "\n".join(lines))

    def show_multistage_success_gate_report(self) -> None:
        r = TKMultistageSuccessGateLab.RESULT
        base = r["baseline_23"]
        new = r["multistage_success_gate_23"]
        weights = r["protocol"]["selected_stage_weights"]
        lines = [
            "太炅 Multistage Success Gate Lab v9.0", "",
            f"학습: {r['protocol']['training']}",
            f"조정검증: {r['protocol']['validation']}",
            f"완전 보류: {r['protocol']['holdout']}",
            f"게이트 피처: {r['protocol']['gate_features']}개",
            f"4개·5개·6개 가중치: {weights['four_plus']} · {weights['five_plus']} · {weights['six_all']}",
            f"선택 임계값: {r['protocol']['selected_gate_threshold']:.3f}", "",
            "[기존 후보23]",
            f"평균 포함: {base['average_hits']:.3f}개",
            f"4개 이상: {base['four_plus_rate']*100:.1f}%",
            f"5개 이상: {base['five_plus_rate']*100:.1f}%",
            f"6개 전부: {base['six_all_rate']*100:.2f}% ({base['six_all_cases']}회)", "",
            "[다단계 성공게이트]",
            f"교체 적용 회차: {new['changed_rounds']}회",
            f"평균 포함: {new['average_hits']:.3f}개",
            f"4개 이상: {new['four_plus_rate']*100:.1f}%",
            f"5개 이상: {new['five_plus_rate']*100:.1f}%",
            f"6개 전부: {new['six_all_rate']*100:.2f}% ({new['six_all_cases']}회)", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"], "",
            f"현재 챔피언: {r['decision']['champion']}",
        ]
        QMessageBox.information(self, "다단계 성공게이트 검증", "\n".join(lines))

    def show_success_dual_gate_report(self) -> None:
        r = TKSuccessDualGateLab.RESULT
        base = r["baseline_23"]
        new = r["success_dual_gate_23"]
        lines = [
            "太炅 Success Dual Gate Lab v8.9", "",
            f"학습: {r['protocol']['training']}",
            f"조정검증: {r['protocol']['validation']}",
            f"완전 보류: {r['protocol']['holdout']}",
            f"게이트 피처: {r['protocol']['gate_features']}개",
            f"성공회차 가중치: {r['protocol']['success_round_weight']:.1f}",
            f"선택 임계값: {r['protocol']['selected_gate_threshold']:.3f}", "",
            "[기존 후보23]",
            f"평균 포함: {base['average_hits']:.3f}개",
            f"4개 이상: {base['four_plus_rate']*100:.1f}%",
            f"5개 이상: {base['five_plus_rate']*100:.1f}%",
            f"6개 전부: {base['six_all_rate']*100:.2f}% ({base['six_all_cases']}회)", "",
            "[성공회차 이중게이트]",
            f"교체 적용 회차: {new['changed_rounds']}회",
            f"평균 포함: {new['average_hits']:.3f}개",
            f"4개 이상: {new['four_plus_rate']*100:.1f}%",
            f"5개 이상: {new['five_plus_rate']*100:.1f}%",
            f"6개 전부: {new['six_all_rate']*100:.2f}% ({new['six_all_cases']}회)", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"], "",
            f"현재 챔피언: {r['decision']['champion']}",
        ]
        QMessageBox.information(self, "성공회차 이중게이트 검증", "\n".join(lines))

    def show_boundary_group_vote_report(self) -> None:
        r = TKBoundaryGroupVoteLab.RESULT
        base = r["baseline_23"]
        new = r["group_vote_23"]
        lines = [
            "太炅 Boundary Group Vote Lab v8.8", "",
            f"학습: {r['protocol']['training']}",
            f"조정검증: {r['protocol']['validation']}",
            f"완전 보류: {r['protocol']['holdout']}",
            f"경계군 크기: {r['protocol']['group_size']}개",
            f"최소 투표점수: {r['protocol']['selected_min_vote_score']} · 임계값 {r['protocol']['selected_threshold']:.4f}", "",
            "[기존 후보23]",
            f"평균 포함: {base['average_hits']:.3f}개",
            f"4개 이상: {base['four_plus_rate']*100:.1f}%",
            f"5개 이상: {base['five_plus_rate']*100:.1f}%",
            f"6개 전부: {base['six_all_rate']*100:.2f}% ({base['six_all_cases']}회)", "",
            "[경계군 집합투표]",
            f"교체 적용 회차: {new['changed_rounds']}회",
            f"평균 포함: {new['average_hits']:.3f}개",
            f"4개 이상: {new['four_plus_rate']*100:.1f}%",
            f"5개 이상: {new['five_plus_rate']*100:.1f}%",
            f"6개 전부: {new['six_all_rate']*100:.2f}% ({new['six_all_cases']}회)", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
            "",
            f"현재 챔피언: {r['decision']['champion']}",
        ]
        QMessageBox.information(self, "경계군 집합투표 검증", "\n".join(lines))

    def show_multi_expert_consensus_report(self) -> None:
        r = TKMultiExpertConsensusLab.RESULT
        base = r["baseline_23"]
        new = r["consensus_swap_23"]
        latest = r["latest"]
        lines = [
            "太炅 Multi-Expert Consensus Lab v8.7", "",
            f"학습: {r['protocol']['training']}",
            f"조정검증: {r['protocol']['validation']}",
            f"완전 보류: {r['protocol']['holdout']}",
            f"전문가: {' · '.join(r['protocol']['experts'])}",
            f"최소 합의: {r['protocol']['selected_min_votes']}명 · 임계값 {r['protocol']['selected_threshold']:.4f}", "",
            "[기존 후보23]",
            f"평균 포함: {base['average_hits']:.3f}개",
            f"4개 이상: {base['four_plus_rate']*100:.1f}%",
            f"5개 이상: {base['five_plus_rate']*100:.1f}%",
            f"6개 전부: {base['six_all_rate']*100:.2f}% ({base['six_all_cases']}회)", "",
            "[다중전문가 합의 교체]",
            f"교체 적용 회차: {new['changed_rounds']}회",
            f"평균 포함: {new['average_hits']:.3f}개",
            f"4개 이상: {new['four_plus_rate']*100:.1f}%",
            f"5개 이상: {new['five_plus_rate']*100:.1f}%",
            f"6개 전부: {new['six_all_rate']*100:.2f}% ({new['six_all_cases']}회)", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"], "",
            f"최신 합의 교체: {latest['consensus_action'] if latest['consensus_action'] else '없음'}",
            "",
            f"현재 챔피언: {r['decision']['champion']}",
        ]
        QMessageBox.information(self, "다중전문가 합의 검증", "\n".join(lines))

    def show_boundary_swap_report(self) -> None:
        r = TKBoundarySwapLab.RESULT
        base = r["baseline_23"]
        swap = r["boundary_swap_23"]
        latest = r["latest"]
        lines = [
            "太炅 Boundary Swap Lab v8.6", "",
            f"학습: {r['protocol']['training']}",
            f"조정검증: {r['protocol']['validation']}",
            f"완전 보류: {r['protocol']['holdout']}",
            f"모델: {r['protocol']['model']} · 피처 {r['protocol']['features']}개", "",
            "[후보23 고정]",
            f"평균 포함: {base['average_hits']:.3f}개",
            f"4개 이상: {base['four_plus_rate']*100:.1f}%",
            f"5개 이상: {base['five_plus_rate']*100:.1f}%",
            f"6개 전부: {base['six_all_rate']*100:.2f}% ({base['six_all_cases']}회)", "",
            "[경계번호 선택 교체]",
            f"교체 적용 회차: {swap['changed_rounds']}회",
            f"평균 포함: {swap['average_hits']:.3f}개",
            f"4개 이상: {swap['four_plus_rate']*100:.1f}%",
            f"5개 이상: {swap['five_plus_rate']*100:.1f}%",
            f"6개 전부: {swap['six_all_rate']*100:.2f}% ({swap['six_all_cases']}회)", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"], "",
            f"최신 제거 후보: {latest['remove_candidate']}",
            f"최신 추가 후보: {latest['add_candidate']}",
            f"교체 점수차: {latest['score_delta']:.4f} · 임계값 {latest['threshold']:.4f}",
            f"최신 교체 적용: {'예' if latest['swap_applied'] else '아니오'}",
            "",
            "현재 챔피언은 v8.3 이중후보 23개를 유지합니다.",
        ]
        QMessageBox.information(self, "경계번호 교체 검증", "\n".join(lines))

    def show_dynamic_candidate_size_report(self) -> None:
        r = TKDynamicCandidateSizeLab.RESULT
        base = r["baseline_23"]
        dyn = r["dynamic_22_23"]
        lines = [
            "太炅 Dynamic Candidate Size Lab v8.5", "",
            f"학습: {r['protocol']['training']}",
            f"조정검증: {r['protocol']['validation']}",
            f"완전 보류: {r['protocol']['holdout']}",
            f"모델: {r['protocol']['model']} · 피처 {r['protocol']['features']}개", "",
            "[후보23 고정]",
            f"평균 후보 수: {base['average_candidate_size']:.2f}개",
            f"평균 포함: {base['average_hits']:.3f}개",
            f"4개 이상: {base['four_plus_rate']*100:.1f}%",
            f"5개 이상: {base['five_plus_rate']*100:.1f}%",
            f"6개 전부: {base['six_all_rate']*100:.2f}% ({base['six_all_cases']}회)", "",
            "[동적 22·23]",
            f"평균 후보 수: {dyn['average_candidate_size']:.2f}개",
            f"23개 선택: {dyn['rounds_using_23']}회 · 22개 선택: {dyn['rounds_using_22']}회",
            f"평균 포함: {dyn['average_hits']:.3f}개",
            f"4개 이상: {dyn['four_plus_rate']*100:.1f}%",
            f"5개 이상: {dyn['five_plus_rate']*100:.1f}%",
            f"6개 전부: {dyn['six_all_rate']*100:.2f}% ({dyn['six_all_cases']}회)", "",
            f"판정: {r['decision']['status']}",
            r["decision"]["reason"],
            "",
            "현재 챔피언은 v8.3 이중후보 23개를 유지합니다.",
        ]
        QMessageBox.information(self, "동적22·23 검증", "\n".join(lines))

    def show_candidate22_report(self) -> None:
        result = TKCandidate22ShrinkLab.RESULT
        before = result["baseline_23"]
        after = result["challenger_22"]
        lines = [
            "太炅 Candidate 22 Shrink Lab v8.4",
            "",
            f"학습: {result['protocol']['training']}",
            f"조정검증: {result['protocol']['validation']}",
            f"완전 보류: {result['protocol']['holdout']}",
            f"시험 설정: {result['protocol']['search_space']:,}개",
            "",
            "[23개 챔피언 → 22개 도전자]",
            f"평균 포함: {before['average_hits']:.3f} → {after['average_hits']:.3f}",
            f"3개 이상: {before['three_plus_rate']*100:.1f}% → {after['three_plus_rate']*100:.1f}%",
            f"4개 이상: {before['four_plus_rate']*100:.1f}% → {after['four_plus_rate']*100:.1f}%",
            f"5개 이상: {before['five_plus_rate']*100:.1f}% → {after['five_plus_rate']*100:.1f}%",
            f"6개 전부: {before['six_all_rate']*100:.2f}% ({before['six_all_cases']}회) → "
            f"{after['six_all_rate']*100:.2f}% ({after['six_all_cases']}회)",
            "",
            f"판정: {result['decision']['status']}",
            result["decision"]["reason"],
            "",
            f"현재 유지: {result['decision']['champion']}",
            f"다음 연구: {result['decision']['next_research']}",
        ]
        QMessageBox.information(self, "후보22 축소검증", "\n".join(lines))

    def show_dual_candidate_report(self) -> None:
        result = TKDualCandidateLab.RESULT
        before = result["baseline_23"]
        after = result["hybrid_23"]
        change = result["change"]
        lines = [
            "太炅 Dual Candidate Lab v8.3",
            "",
            f"학습: {result['protocol']['training']}",
            f"조정검증: {result['protocol']['validation']}",
            f"완전 보류: {result['protocol']['holdout']}",
            f"시험 가중치: {result['protocol']['search_count']:,}개",
            "",
            "[구조]",
            "기존 핵심후보 20개를 유지",
            "공격형 보조엔진에서 비중복 번호 3개 추가",
            "최종 후보 23개",
            "",
            "[보류검증 비교]",
            f"기존23 평균: {before['avg']:.3f}개",
            f"이중23 평균: {after['avg']:.3f}개",
            f"기존 4개 이상: {before['r4']*100:.1f}%",
            f"이중 4개 이상: {after['r4']*100:.1f}%",
            f"기존 5개 이상: {before['r5']*100:.1f}%",
            f"이중 5개 이상: {after['r5']*100:.1f}%",
            f"기존 6개 전부: {before['r6']*100:.2f}% ({change['six_all_cases_before']}회)",
            f"이중 6개 전부: {after['r6']*100:.2f}% ({change['six_all_cases_after']}회)",
            "",
            f"판정: {result['decision']['status']} · {result['decision']['role']}",
            result["decision"]["reason"],
            "",
            f"현재 핵심20: {' · '.join(map(str, result['latest']['core20']))}",
            f"보조추가3: {' · '.join(map(str, result['latest']['support_add3']))}",
            f"최종후보23: {' · '.join(map(str, result['latest']['hybrid23']))}",
            "",
            "※ 기존 안정형 챔피언은 유지하며, 이중후보 엔진은 6개 전부 포함을 노리는 공격형 보조엔진입니다.",
        ]
        QMessageBox.information(self, "이중후보 검증", "\n".join(lines))

    def show_candidate_shrink_report(self) -> None:
        result = TKCandidateShrinkLab.RESULT
        lines = [
            "太炅 Candidate Shrink Lab v8.2",
            "",
            f"학습: {result['protocol']['training']}",
            f"조정검증: {result['protocol']['validation']}",
            f"완전 보류: {result['protocol']['holdout']}",
            f"도전자: {result['protocol']['challengers']:,}개",
            "",
            "[기존 챔피언 후보 수별 보류성적]",
        ]
        for size in (15, 16, 17, 18, 19, 20, 21, 22, 25):
            row = result["baseline"][str(size)]
            lines.append(
                f"{size}개 · 평균 {row['avg']:.3f} · "
                f"5개+ {row['r5']*100:.1f}% · "
                f"6개 전부 {row['r6']*100:.2f}% · 최고 {row['max']}개"
            )
        lines.extend([
            "",
            f"6개 전부가 처음 확인된 최소 후보 수: {result['decision']['smallest_candidate_with_six_all']}개",
            f"판정: {result['decision']['challenger_status']}",
            result["decision"]["reason"],
            "",
            "현재 챔피언은 교체하지 않습니다.",
            "다음 연구는 19~22개에서 6개 전부 포함률을 유지한 뒤 한 개씩 줄이는 방식입니다.",
        ])
        QMessageBox.information(self, "후보축소 검증", "\n".join(lines))

    def show_candidate_lab_report(self) -> None:
        result = TKCandidateLab.RESULT
        lines = [
            "太炅 AI Candidate Lab 중간결과",
            "",
            f"초기학습: {result['method']['initial_training']}",
            f"성장검증: {result['method']['growth']}",
            f"완전 보류검증: {result['method']['holdout']}",
            "",
            "[후보 개수별 보류검증]",
        ]
        for size in (10, 12, 15, 18, 20):
            row = result["holdout_results"][str(size)]
            lines.append(
                f"{size}개 · 평균 {row['average_hits']:.3f}개 · "
                f"4개 이상 {row['four_plus_rate']*100:.1f}% · "
                f"5개 이상 {row['five_plus_rate']*100:.1f}% · "
                f"최고 {row['max_hits']}개"
            )
        selected = result["holdout_results"]["15"]
        lines.extend([
            "",
            "[현재 적용: 후보 15개]",
            f"평균 포함: {selected['average_hits']:.3f}개",
            f"3개 이상: {selected['three_plus_rate']*100:.1f}%",
            f"4개 이상: {selected['four_plus_rate']*100:.1f}%",
            f"5개 이상: {selected['five_plus_rate']*100:.1f}%",
            f"6개 전부: {selected['six_hits_rate']*100:.2f}%",
            f"최고: {selected['max_hits']}개",
            "",
            f"현재 최신 후보15: {' · '.join(map(str, result['latest_top15']))}",
            "",
            "※ 후보번호 적중과 조합 배치를 분리하기 위한 중간 연구판입니다.",
            "※ 과거 검증 결과는 다음 회차 당첨을 보장하지 않습니다.",
        ])
        QMessageBox.information(self, "후보15 연구결과", "\n".join(lines))

    def show_lab_report(self) -> None:
        r = TKV8EvolutionLab.RESULT
        c = r["champion"]; e = r["ensemble"]
        lines = [
            "太炅 AI Evolution Lab v8.0", "",
            f"초기학습: {r['data']['initial_training']}",
            f"성장구간: {r['data']['growth']}",
            f"완전 보류검증: {r['data']['holdout']}", "",
            "[AI 엔진 월드컵]",
            f"엔진 수: {r['world_cup']['population']}개",
            f"세대 수: {r['world_cup']['generations']}세대",
            f"총 엔진 평가: {r['world_cup']['total_engine_evaluations']:,}회", "",
            "[최종 챔피언]",
            f"보류 TOP15 평균: {c['holdout_avg']:.3f}개",
            f"3개 이상: {c['holdout_3plus']*100:.1f}%",
            f"4개 이상: {c['holdout_4plus']*100:.1f}%",
            f"5개 이상: {c['holdout_5plus']*100:.1f}%",
            f"최고: {c['holdout_max']}개", "",
            "[메타 앙상블]",
            f"보류 TOP15 평균: {e['holdout_avg']:.3f}개",
            f"4개 이상: {e['holdout_4plus']*100:.1f}%",
            f"5개 이상: {e['holdout_5plus']*100:.1f}%",
            f"최고: {e['holdout_max']}개", "",
            f"다음 추천 후보 TOP20: {' · '.join(map(str, r['next_top20']))}", "",
            "※ 1133~1218회는 학습에 사용하지 않은 보류구간입니다.",
            "※ 과거 검증은 미래 당첨을 보장하지 않습니다.",
        ]
        QMessageBox.information(self, "AI 연구소 결과", "\n".join(lines))

    def show_evolution_report(self) -> None:
        result = TKEvolutionEngine.RESULT
        base = result["holdout_base"]
        evolved = result["holdout_evolved"]
        lines = [
            "TK AI 회차별 성장결과",
            "",
            f"초기 학습: {result['initial_training']}",
            f"성장 시뮬레이션: {result['growth_range']}",
            f"완전 보류검증: {result['holdout_range']}",
            f"배포용 재성장: {result['deployment_growth_range']}",
            f"성장 단계: {result['deployment_steps']}회",
            f"가중치 채택 횟수: {result['adoptions_deployment']}회",
            "",
            "[보류구간 비교]",
            f"기존 TOP15 평균: {base['avg']:.3f}개",
            f"성장 TOP15 평균: {evolved['avg']:.3f}개",
            f"기존 4개 이상: {base['r4'] * 100:.1f}%",
            f"성장 4개 이상: {evolved['r4'] * 100:.1f}%",
            f"기존 5개 이상: {base['r5'] * 100:.1f}%",
            f"성장 5개 이상: {evolved['r5'] * 100:.1f}%",
            f"기존 최고: {base['max']}개",
            f"성장 최고: {evolved['max']}개",
            "",
            "※ 1133~1218회는 성장 과정에 넣지 않은 별도 검증 결과입니다.",
            "※ 배포용 엔진은 검증 후 같은 규칙으로 1218회까지 다시 성장시켰습니다.",
            "※ 과거 성적은 다음 회차 당첨을 보장하지 않습니다.",
        ]
        QMessageBox.information(self, "AI 성장결과", "\\n".join(lines))

    def show_self_check_report(self) -> None:
        path = Path(__file__).resolve().parent / "자동테스트_결과.json"
        try:
            report = json.loads(path.read_text(encoding="utf-8"))
            lines = [f"전체 {report['summary']['total']}개 · 통과 {report['summary']['passed']}개 · 실패 {report['summary']['failed']}개", ""]
            for item in report["tests"]:
                lines.append(f"[{'통과' if item['passed'] else '실패'}] {item['name']} · {item['detail']}")
            QMessageBox.information(self, "내부 자동검증 결과", "\n".join(lines))
        except Exception as exc:
            QMessageBox.warning(self, "내부 자동검증 결과", str(exc))

    def show_performance_report(self) -> None:
        common = TKPerformanceEngine.OPTIMIZATION_RESULT
        category = self.rec_category.currentText()
        if category in ("자체추천", "특이패턴추천"):
            mode = self.pattern_mode_combo.currentText()
        elif category == "통합데이터추천":
            mode = self.mixed_preset_combo.currentText()
        else:
            mode = "기본"
        selected = TKEngineAudit.find(category, mode)
        lines = [
            "TK 최적화·항목별 검증 결과", "",
            "[공통 성과엔진 30,000개 설정 경쟁]",
            f"시험 설정: {common['tested_settings']:,}개",
            f"최종 보류: {common['holdout']['round_start']}~{common['holdout']['round_end']}회",
            f"TOP15 평균 포함: {common['holdout']['average_top15_hits']:.3f}개", "",
            f"[현재 선택: {category} / {mode}]",
        ]
        if selected:
            lines += [
                f"검증 회차: {selected['rounds']}회",
                f"TOP15 평균 포함: {selected['top15_average_hits']:.3f}개",
                f"TOP15 3개 이상: {selected['top15_3plus_rate']*100:.1f}%",
                f"TOP15 4개 이상: {selected['top15_4plus_rate']*100:.1f}%",
                f"추천100조합 평균 최고적중: {selected['combo100_average_best_hits']:.3f}개",
                f"추천100조합 4개 이상: {selected['combo100_4plus_rate']*100:.1f}%",
                f"추천100조합 5개 이상: {selected['combo100_5plus_rate']*100:.1f}%",
                f"추천100조합 최고: {selected['combo100_max_hits']}개",
            ]
        lines += ["", "[항목별 종합순위 TOP10]"]
        for i,row in enumerate(TKEngineAudit.rows()[:10],1):
            lines.append(f"{i}. {row['category']} / {row['mode']} · TOP15 {row['top15_average_hits']:.2f} · 100조합 {row['combo100_average_best_hits']:.2f}")
        lines += ["", "※ 통합데이터추천은 직전 5회 번호를 대체 입력으로 사용한 비교입니다.", "※ 과거 검증은 미래 당첨을 보장하지 않습니다."]
        QMessageBox.information(self, "최적화·항목별 검증 결과", "\n".join(lines))

    def _draw_data_path(self) -> Path:
        root = Path(os.getenv("LOCALAPPDATA", Path.home())) / "TaegyeongLottoLab"
        root.mkdir(parents=True, exist_ok=True)
        return root / "manual_draws.json"

    def _backup_draw_data(self) -> None:
        path = self._draw_data_path()
        if not path.exists():
            return
        backup_dir = path.parent / "backup"
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(path, backup_dir / f"manual_draws_{stamp}.json")

    def manual_add_draw(self) -> None:
        if not self.analyzer.draws:
            QMessageBox.information(self, "데이터 필요", "먼저 역대 Excel을 불러오세요.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("수동 회차 추가")
        form = QFormLayout(dialog)
        round_input = QLineEdit(str(self.analyzer.draws[-1].round_no + 1))
        numbers_input = QLineEdit()
        numbers_input.setPlaceholderText("예: 3 11 17 24 33 42")
        bonus_input = QLineEdit()
        bonus_input.setPlaceholderText("예: 7")
        form.addRow("회차", round_input)
        form.addRow("당첨번호 6개", numbers_input)
        form.addRow("보너스번호", bonus_input)
        save_btn = QPushButton("검사 후 저장")
        form.addRow(save_btn)

        def save():
            try:
                round_no = int(round_input.text().strip())
                nums = sorted(parse_numbers(numbers_input.text()))
                bonus_values = parse_numbers(bonus_input.text())
                if len(nums) != 6 or len(set(nums)) != 6:
                    raise ValueError("서로 다른 당첨번호 6개를 입력하세요.")
                if len(bonus_values) != 1:
                    raise ValueError("보너스번호 1개를 입력하세요.")
                bonus = bonus_values[0]
                if bonus in nums:
                    raise ValueError("보너스번호가 당첨번호와 중복됩니다.")
                existing = {draw.round_no for draw in self.analyzer.draws}
                if round_no in existing:
                    raise ValueError("이미 등록된 회차입니다.")
                expected = self.analyzer.draws[-1].round_no + 1
                if round_no != expected:
                    answer = QMessageBox.question(
                        dialog, "회차 확인",
                        f"다음 예상 회차는 {expected}회입니다. {round_no}회로 저장할까요?"
                    )
                    if answer != QMessageBox.Yes:
                        return
                self._backup_draw_data()
                path = self._draw_data_path()
                records = []
                if path.exists():
                    records = json.loads(path.read_text(encoding="utf-8"))
                records.append({
                    "round_no": round_no,
                    "numbers": nums,
                    "bonus": bonus,
                    "saved_at": datetime.now().isoformat(timespec="seconds"),
                })
                path.write_text(
                    json.dumps(records, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
                self.analyzer.draws.append(
                    Draw(round_no, tuple(nums), bonus)
                )
                self.analyzer.draws.sort(key=lambda draw: draw.round_no)
                self.analyzer._analyze()
                self.pattern_cache.clear()
                dialog.accept()
                self.refresh_all()
                self.save_session()
                QMessageBox.information(
                    self, "저장 완료",
                    f"{round_no}회가 추가되었습니다. 분석 캐시도 새 데이터 기준으로 갱신했습니다."
                )
            except Exception as exc:
                QMessageBox.warning(dialog, "입력 오류", str(exc))

        save_btn.clicked.connect(save)
        dialog.exec()

    @staticmethod
    def _parse_official_latest_html(html: str):
        rounds = [int(value) for value in re.findall(r"(\d{1,4})회", html)]
        if not rounds:
            raise ValueError("공식 페이지에서 최신 회차를 찾지 못했습니다.")
        latest = max(rounds)
        # 공식 결과 페이지의 최신 회차 주변에서 번호 6개와 보너스를 탐색
        marker = html.find(f"{latest}회")
        segment = html[marker:marker + 12000] if marker >= 0 else html[:12000]
        clean = re.sub(r"<[^>]+>", " ", segment)
        values = [int(v) for v in re.findall(r"(?<!\d)([1-9]|[1-3]\d|4[0-5])(?!\d)", clean)]
        # 중복을 보존하되 첫 정상적인 7개 연속 범위를 탐색
        for i in range(max(0, len(values) - 30)):
            candidate = values[i:i + 7]
            if len(candidate) == 7 and len(set(candidate[:6])) == 6 and candidate[6] not in candidate[:6]:
                return latest, sorted(candidate[:6]), candidate[6]
        raise ValueError("공식 페이지에서 번호를 안전하게 해석하지 못했습니다.")

    def check_latest_draw(self) -> None:
        if not self.analyzer.draws:
            QMessageBox.information(
                self, "데이터 필요", "먼저 역대 Excel을 불러오세요."
            )
            return
        try:
            request = urllib.request.Request(
                "https://www.dhlottery.co.kr/lt645/result",
                headers={"User-Agent": "Mozilla/5.0"},
            )
            with urllib.request.urlopen(request, timeout=15) as response:
                html = response.read().decode("utf-8", errors="ignore")

            latest, numbers, bonus = self._parse_official_latest_html(html)
            current = self.analyzer.draws[-1].round_no
            if latest <= current:
                QMessageBox.information(
                    self, "최신 회차 확인",
                    f"현재 보유 데이터 {current}회가 최신입니다."
                )
                return

            answer = QMessageBox.question(
                self,
                "새 회차 발견",
                f"현재 보유: {current}회\n"
                f"공식 최신: {latest}회\n"
                f"당첨번호: {' · '.join(map(str, numbers))}\n"
                f"보너스번호: {bonus}\n\n"
                "확인한 번호를 바로 업데이트할까요?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if answer != QMessageBox.Yes:
                return

            if latest in {draw.round_no for draw in self.analyzer.draws}:
                QMessageBox.information(
                    self, "중복 회차", f"{latest}회는 이미 등록되어 있습니다."
                )
                return

            expected = current + 1
            if latest != expected:
                confirm = QMessageBox.question(
                    self,
                    "회차 누락 경고",
                    f"다음 예상 회차는 {expected}회지만 최신 회차는 {latest}회입니다.\n"
                    "그래도 추가할까요?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if confirm != QMessageBox.Yes:
                    return

            self._backup_draw_data()
            path = self._draw_data_path()
            records = []
            if path.exists():
                try:
                    records = json.loads(path.read_text(encoding="utf-8"))
                except Exception:
                    records = []

            if not any(int(r.get("round_no", -1)) == latest for r in records):
                records.append({
                    "round_no": latest,
                    "numbers": numbers,
                    "bonus": bonus,
                    "source": "official_auto_check",
                    "saved_at": datetime.now().isoformat(timespec="seconds"),
                })
                path.write_text(
                    json.dumps(records, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )

            self.analyzer.draws.append(Draw(latest, tuple(numbers), bonus))
            self.analyzer.draws.sort(key=lambda draw: draw.round_no)
            self.analyzer._analyze()
            self.pattern_cache.clear()
            self.recommendations = []
            self.update_source_counts()
            self.save_session()

            QMessageBox.information(
                self,
                "업데이트 완료",
                f"{latest}회가 저장되었습니다.\n"
                f"당첨번호: {' · '.join(map(str, numbers))}\n"
                f"보너스번호: {bonus}\n\n"
                "전체 통계와 추천엔진이 새 회차 기준으로 갱신되었습니다."
            )
            self.generate_recommendations()

        except Exception as exc:
            QMessageBox.warning(
                self,
                "자동 확인 실패",
                "공식 페이지 구조 변경 또는 인터넷 연결 문제로 자동 확인하지 못했습니다.\n"
                f"상세: {exc}\n\n"
                "데이터 관리 > 수동 회차 추가를 이용하세요."
            )

    def show_pattern_briefing(self) -> None:
        if not self.analyzer.draws:
            QMessageBox.information(
                self, "데이터 필요", "먼저 역대 Excel을 불러오세요."
            )
            return
        try:
            recommender = Recommender(self.analyzer)
            cache_key = f"briefing:{len(self.analyzer.draws)}"
            cached = self.pattern_cache.get(cache_key)
            if cached is None:
                briefing = recommender.pattern_briefing()
                board, reliability = recommender.pattern_board()
                self.pattern_cache[cache_key] = (briefing, board, reliability)
            else:
                briefing, board, reliability = cached
            lines = [briefing, "", "번호별 패턴투표 TOP20"]
            for rank, row in enumerate(board[:20], 1):
                patterns = ", ".join(row["patterns"]) or "단독 점수"
                reasons = " / ".join(row["reasons"][:2])
                lines.append(
                    f"{rank:02d}. {row['number']}번 | {row['votes']}표 | "
                    f"{row['score']:.1f}점 | {patterns}"
                    + (f" | {reasons}" if reasons else "")
                )
            self.detail_box.setPlainText("\n".join(lines))
            self.rec_status.setText(
                "이번 주 패턴 브리핑 완료 · 상세창에서 핵심번호와 제외번호를 확인하세요."
            )
        except Exception as exc:
            QMessageBox.critical(self, "패턴 분석 오류", str(exc))

    def search_rounds(self) -> None:
        if not self.analyzer.draws:
            QMessageBox.information(
                self, "데이터 필요", "먼저 역대 Excel을 불러오세요."
            )
            return
        try:
            values = parse_numbers(self.round_search_input.text())
            if not values:
                raise ValueError("검색할 번호를 입력하세요.")
            target = set(values)
            rows = []
            for draw in reversed(self.analyzer.draws):
                matched = sorted(target & set(draw.numbers))
                if matched:
                    rows.append(
                        f"{draw.round_no}회 | {' · '.join(map(str, draw.numbers))} | "
                        f"일치 {len(matched)}개: {', '.join(map(str, matched))}"
                    )
                if len(rows) >= 100:
                    break
            if not rows:
                rows = ["일치하는 회차가 없습니다."]
            self.detail_box.setPlainText(
                f"번호 검색: {', '.join(map(str, values))}\n\n"
                + "\n".join(rows)
            )
        except Exception as exc:
            QMessageBox.warning(self, "회차 검색", str(exc))

    def run_strategy_battle(self) -> None:
        if len(self.analyzer.draws) < 120:
            QMessageBox.information(
                self,
                "데이터 부족",
                "전략 배틀은 최소 120회 이상의 역대 데이터가 필요합니다.",
            )
            return

        answer = QMessageBox.question(
            self,
            "전략 배틀",
            "최근 100회를 과거 데이터만 사용해 백테스트합니다.\n"
            "PC 성능에 따라 1~3분 정도 걸릴 수 있습니다. 실행할까요?",
        )
        if answer != QMessageBox.Yes:
            return

        self.strategy_battle_btn.setEnabled(False)
        self.rec_status.setText("전략 배틀 준비 중...")
        QApplication.processEvents()

        try:
            strategies = list(Recommender.STRATEGY_WEIGHTS)
            stats = {
                name: {
                    "three_plus": 0,
                    "four_plus": 0,
                    "five_plus": 0,
                    "six": 0,
                    "hit_sum": 0,
                    "rounds": 0,
                }
                for name in strategies
            }

            draws = self.analyzer.draws
            start_index = max(20, len(draws) - 100)

            for test_no, target_index in enumerate(
                range(start_index, len(draws)), 1
            ):
                history = draws[:target_index]
                target = set(draws[target_index].numbers)

                temp = LottoAnalyzer()
                temp.draws = list(history)
                temp._analyze()
                recommender = Recommender(temp)

                # 과거 출현빈도 상위 15개를 입력번호로 가정해 전략별 TOP10 생성
                ranked_numbers = [
                    n for n, _ in temp.number_counts.most_common(15)
                ]
                source_weights = Counter({
                    n: max(1, temp.number_counts[n])
                    for n in ranked_numbers
                })

                for strategy in strategies:
                    rows = recommender.generate(
                        source_weights,
                        10,
                        20,
                        300,
                        True,
                        "추천조합",
                        strategy=strategy,
                    )
                    best_hit = max(
                        (len(target & set(combo)) for _, combo, _ in rows),
                        default=0,
                    )
                    item = stats[strategy]
                    item["rounds"] += 1
                    item["hit_sum"] += best_hit
                    if best_hit >= 3:
                        item["three_plus"] += 1
                    if best_hit >= 4:
                        item["four_plus"] += 1
                    if best_hit >= 5:
                        item["five_plus"] += 1
                    if best_hit >= 6:
                        item["six"] += 1

                if test_no % 5 == 0:
                    self.rec_status.setText(
                        f"전략 배틀 진행 중 · {test_no}/100회"
                    )
                    QApplication.processEvents()

            rows = []
            for strategy, item in stats.items():
                rounds = max(1, item["rounds"])
                average = item["hit_sum"] / rounds
                battle_score = (
                    item["three_plus"] * 1
                    + item["four_plus"] * 3
                    + item["five_plus"] * 10
                    + item["six"] * 50
                    + average * 10
                )
                rows.append((battle_score, strategy, item, average))

            rows.sort(reverse=True)

            lines = [
                "최근 100회 전략 배틀 결과",
                "각 회차마다 이전 회차 데이터만 사용하고 전략별 TOP10 중 최고 적중을 비교합니다.",
                "",
            ]
            for rank, (_, strategy, item, average) in enumerate(rows, 1):
                lines.append(
                    f"{rank}위 {strategy} | 평균 최고적중 {average:.2f}개 | "
                    f"3개+ {item['three_plus']}회 | 4개+ {item['four_plus']}회 | "
                    f"5개+ {item['five_plus']}회 | 6개 {item['six']}회"
                )

            result_text = "\n".join(lines)
            self.detail_box.setPlainText(result_text)
            self.rec_status.setText(
                f"전략 배틀 완료 · 1위: {rows[0][1]}"
            )
            QMessageBox.information(
                self,
                "전략 배틀 완료",
                f"최근 100회 기준 1위 전략은 '{rows[0][1]}'입니다.\n"
                "자세한 결과는 추천 결과 아래 상세창에서 확인하세요.",
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "전략 배틀 오류",
                f"{exc}\n\n{traceback.format_exc(limit=3)}",
            )
        finally:
            self.strategy_battle_btn.setEnabled(True)

    def show_recommend_detail(self, row: int, _column: int) -> None:
        if row < 0 or row >= len(self.recommendations):
            return

        recommender = Recommender(self.analyzer)
        score, combo, metrics = self.recommendations[row]
        category = self.rec_category.currentText()

        if category == "자체추천":
            fixed_numbers = ()
            candidate_numbers = ()
        else:
            fixed_numbers = self.fixed_numbers()
            candidate_numbers = self.candidate_numbers()

        confidence = recommender.confidence_score(score, metrics)
        grade = recommender.confidence_grade(confidence)
        reason = recommender.recommendation_reason(
            metrics, fixed_numbers, candidate_numbers, combo
        )

        pair_lines = [
            f"{a}↔{b}: {count}회"
            for (a, b), count in recommender.pair_details(combo, 5)
        ]
        similar_lines = [
            f"{round_no}회 · 유사도 {similarity:.1f}% · "
            + " · ".join(map(str, numbers))
            for round_no, similarity, numbers
            in recommender.historical_similar_draws(combo, 5)
        ]
        triple_lines = [
            f"{a}-{b}-{c}: {count}회"
            for (a, b, c), count in recommender.triple_details(combo, 5)
        ]

        mixed_detail = ""
        if metrics.get("mixed_preset"):
            mixed_detail = (
                f"통합비율: {metrics['mixed_preset']} "
                f"(입력 {metrics['mixed_input_weight']:.0%} / "
                f"100회 {metrics['mixed_100_weight']:.0%} / "
                f"500회 {metrics['mixed_500_weight']:.0%} / "
                f"1000회 {metrics['mixed_1000_weight']:.0%})\n"
                f"구간점수: 100회 {metrics['score100']:.1f} / "
                f"500회 {metrics['score500']:.1f} / "
                f"1000회 {metrics['score1000']:.1f}\n"
            )

        text = (
            f"추천조합: {' · '.join(map(str, combo))}\n"
            f"추천신뢰도: {confidence:.1f}점 ({grade}등급)\n"
            f"추천전략: {metrics.get('strategy', '자동')}\n"
            f"추천이유: {reason}\n"
            f"{mixed_detail}"
            f"합계: {sum(combo)} / 홀수 {sum(n % 2 for n in combo)}개 / "
            f"고번호 {sum(n >= 23 for n in combo)}개\n\n"
            f"[동반출현 횟수]\n" + "\n".join(pair_lines) +
            f"\n\n[트리플 출현 횟수]\n" + "\n".join(triple_lines) +
            f"\n\n[세부 점수]\n"
            f"나온횟수 {metrics.get('input', 0):.1f} / "
            f"동반수 {metrics.get('pair', 0):.1f} / "
            f"트리플 {metrics.get('triple', 0):.1f} / "
            f"최근패턴 {metrics.get('recent', 0):.1f} / "
            f"조합균형 {metrics.get('structure', 0):.1f}"
            + (
                f"\n\n[특이패턴]\n"
                f"패턴투표 {metrics.get('pattern_votes', 0)}표 / "
                f"패턴점수 {metrics.get('pattern_score', 0):.1f}\n"
                f"주요패턴: {', '.join(metrics.get('pattern_names', []))}\n"
                f"추천근거: {' / '.join(metrics.get('pattern_reasons', []))}"
                if metrics.get("pattern_names") else ""
            )
            + (
                "\n\n[자체추천 패턴 적용]\n"
                f"현재패턴: {metrics.get('self_pattern_mode', '')}\n"
                f"패턴비중: {metrics.get('self_pattern_mix', 0) * 100:.0f}%\n"
                f"패턴TOP10: {' · '.join(map(str, metrics.get('self_pattern_top10', [])))}\n"
                f"후보번호풀: {' · '.join(map(str, metrics.get('self_candidate_pool', [])))}"
                if metrics.get("self_pattern_mode") else ""
            )
            + (
                "\n\n[성과최적 엔진 근거]\n"
                + "\n".join(metrics.get("performance_reasons", []))
                if metrics.get("performance_reasons") else ""
            )
            + "\n\n[역대 유사 회차]\n" + "\n".join(similar_lines)
        )
        self.detail_box.setPlainText(text)

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
        if category == "자체추천":
            fixed_numbers = ()
            excluded_numbers = ()
            candidate_numbers = ()
        else:
            fixed_numbers = self.fixed_numbers()
            excluded_numbers = self.excluded_numbers()
            candidate_numbers = self.candidate_numbers()

        rows = []
        for rank, (score, combo, metrics) in enumerate(self.recommendations, 1):
            rows.append({
                "카테고리": category,
                "추천전략": metrics.get("strategy", "자체추천"),
                "통합프리셋": metrics.get("mixed_preset", ""),
                "입력비중": metrics.get("mixed_input_weight", ""),
                "최근100회비중": metrics.get("mixed_100_weight", ""),
                "최근500회비중": metrics.get("mixed_500_weight", ""),
                "최근1000회비중": metrics.get("mixed_1000_weight", ""),
                "통합데이터후보풀": ", ".join(map(str, metrics.get("mixed_candidate_pool", []))),
                "통합입력TOP10": ", ".join(map(str, metrics.get("mixed_top_input", []))),
                "통합최근100TOP10": ", ".join(map(str, metrics.get("mixed_top_100", []))),
                "통합최근500TOP10": ", ".join(map(str, metrics.get("mixed_top_500", []))),
                "통합최근1000TOP10": ", ".join(map(str, metrics.get("mixed_top_1000", []))),
                "패턴모드": metrics.get("pattern_mode", ""),
                "패턴투표수": metrics.get("pattern_votes", ""),
                "패턴점수": metrics.get("pattern_score", ""),
                "주요패턴": ", ".join(metrics.get("pattern_names", [])),
                "패턴추천근거": " / ".join(metrics.get("pattern_reasons", [])),
                "자체추천패턴": metrics.get("self_pattern_mode", ""),
                "자체추천패턴점수": metrics.get("self_pattern_score", ""),
                "자체추천패턴비중": metrics.get("self_pattern_mix", ""),
                "자체추천패턴TOP10": ", ".join(map(str, metrics.get("self_pattern_top10", []))),
                "자체추천후보풀": ", ".join(map(str, metrics.get("self_candidate_pool", []))),
                "성과최적점수": metrics.get("performance", ""),
                "성과최적근거": " / ".join(metrics.get("performance_reasons", [])),
                "AI진화버전": metrics.get("evolution_version", ""),
                "AI성장단계": metrics.get("evolution_steps", ""),
                "AI진화점수": metrics.get("evolution", ""),
                "V8연구소엔진": metrics.get("v8_engine", ""),
                "V8총엔진평가": metrics.get("v8_worldcup_evaluations", ""),
                "V8보류평균": metrics.get("v8_holdout_avg", ""),
                "V8보류4개이상": metrics.get("v8_holdout_4plus", ""),
                "후보번호연구": metrics.get("candidate_lab", ""),
                "후보번호풀": ", ".join(map(str, metrics.get("candidate_pool", []))),
                "후보번호개수": metrics.get("candidate_pool_size", ""),
                "후보15보류평균": metrics.get("candidate_holdout_average", ""),
                "후보15보류4개이상": metrics.get("candidate_holdout_4plus", ""),
                "후보15보류5개이상": metrics.get("candidate_holdout_5plus", ""),
                "이중후보연구": metrics.get("dual_candidate_lab", ""),
                "이중후보풀": ", ".join(map(str, metrics.get("candidate_pool", []))),
                "이중보조추가": ", ".join(map(str, metrics.get("support_additions", []))),
                "이중보류평균": metrics.get("holdout_average", ""),
                "이중보류4개이상": metrics.get("holdout_4plus", ""),
                "이중보류5개이상": metrics.get("holdout_5plus", ""),
                "이중보류6개전부": metrics.get("holdout_6all", ""),
                "순위": rank,
                "번호1": combo[0], "번호2": combo[1], "번호3": combo[2],
                "번호4": combo[3], "번호5": combo[4], "번호6": combo[5],
                "필수번호": ", ".join(map(str, fixed_numbers)),
                "제외번호": ", ".join(map(str, excluded_numbers)),
                "후보번호": ", ".join(map(str, candidate_numbers)),
                "추천이유": recommender.recommendation_reason(
                    metrics,
                    fixed_numbers,
                    candidate_numbers,
                    combo,
                ),
                "추천신뢰도": recommender.confidence_score(score, metrics),
                "등급": recommender.confidence_grade(
                    recommender.confidence_score(score, metrics)
                ),
                "카테고리점수": round(score, 1),
                "종합점수": round(metrics["composite"], 1),
                "입력횟수점수": round(metrics["input"], 1),
                "동반수점수": round(metrics["pair"], 1),
                "트리플점수": round(metrics["triple"], 1),
                "최근패턴점수": round(metrics["recent"], 1),
                "자체추천점수": round(metrics.get("self", 0.0), 1),
                "동반출현횟수": ", ".join(
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
