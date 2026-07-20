
from __future__ import annotations
import csv, hashlib, heapq, json, math, os, sys, tempfile, time, traceback
from collections import Counter
from dataclasses import dataclass, asdict
from itertools import combinations
from pathlib import Path
from typing import Iterable

import pandas as pd
from openpyxl import Workbook
from PySide6.QtCore import Qt, QProcess, QProcessEnvironment, QTimer, Signal
from PySide6.QtGui import QAction, QFont
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QFrame, QHBoxLayout, QLabel, QMainWindow,
    QMessageBox, QPushButton, QPlainTextEdit, QProgressBar, QSpinBox,
    QSplitter, QStackedWidget, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget, QFormLayout, QComboBox, QCheckBox
)

APP_NAME = "太炅 Lotto Lab Ultimate"
APP_VERSION = "31.0 Continuous Auto Research Build"
PRICE_PER_LINE = 1000
PAYOUT = {6: 1_500_000_000, "5b": 50_000_000, 5: 1_300_000, 4: 5_000, 3: 1_000}


def runtime_base_dir() -> Path:
    """Return the installed/extracted application directory."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def append_runtime_log(filename: str, message: str) -> None:
    path = runtime_base_dir() / filename
    stamp = time.strftime("%Y-%m-%d %H:%M:%S")
    try:
        with path.open("a", encoding="utf-8") as f:
            f.write(f"\n[{stamp}]\n{message}\n")
    except Exception:
        pass

def install_exception_logger() -> None:
    def handle(exc_type, exc_value, exc_tb):
        detail = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        append_runtime_log("error.log", detail)
        try:
            QMessageBox.critical(
                None,
                "太炅 Lotto Lab 오류",
                "프로그램 오류가 발생했습니다.\n"
                "같은 폴더의 error.log에 원인이 기록되었습니다."
            )
        except Exception:
            pass
    sys.excepthook = handle

@dataclass(frozen=True)
class LottoRow:
    round_no: int
    numbers: tuple[int, int, int, int, int, int]
    bonus: int | None = None

@dataclass
class RoundResult:
    round_no: int
    max_match: int
    candidate_hits: int
    candidate_all6: int
    first: int
    second: int
    third: int
    fourth: int
    fifth: int
    winnings: int
    cost: int

@dataclass
class EngineSummary:
    engine: str
    rounds: int
    roi: float
    first: int
    second: int
    third: int
    fourth: int
    fifth: int
    avg_max_match: float
    max4_plus_rounds: int
    max5_plus_rounds: int
    exact6_rounds: int
    candidate_all6_rounds: int
    avg_candidate_hits: float
    reproducibility_hash: str

class LottoDataParser:
    ROUND_HINTS = ("회차", "회", "round", "draw")
    BONUS_HINTS = ("보너스", "bonus")
    NUMBER_HINTS = ("번호", "당첨", "ball", "num")

    @staticmethod
    def _normalize(value: object) -> str:
        return str(value).strip().lower().replace(" ", "").replace("_", "")

    @classmethod
    def _find_round_column(cls, columns: Iterable[object]):
        for col in columns:
            if any(h in cls._normalize(col) for h in cls.ROUND_HINTS):
                return col
        return None

    @classmethod
    def _find_bonus_column(cls, columns: Iterable[object]):
        for col in columns:
            if any(h in cls._normalize(col) for h in cls.BONUS_HINTS):
                return col
        return None

    @classmethod
    def _find_number_columns(cls, df, round_col, bonus_col):
        excluded = {round_col, bonus_col}
        cols = [c for c in df.columns if c not in excluded and any(h in cls._normalize(c) for h in cls.NUMBER_HINTS)]
        if len(cols) < 6:
            cols = []
            for c in df.columns:
                if c in excluded:
                    continue
                n = pd.to_numeric(df[c], errors="coerce").dropna()
                if not n.empty and n.between(1,45).mean() >= .7:
                    cols.append(c)
        return cols[:6]

    @classmethod
    def parse_excel(cls, file_path):
        path = Path(file_path)
        wb = pd.ExcelFile(path)
        best = []
        for sheet in wb.sheet_names:
            df = pd.read_excel(path, sheet_name=sheet)
            if df.empty:
                continue
            rc = cls._find_round_column(df.columns)
            bc = cls._find_bonus_column(df.columns)
            nc = cls._find_number_columns(df, rc, bc)
            if len(nc) < 6:
                continue
            rows = []
            for i, row in df.iterrows():
                try:
                    nums = [int(float(row[c])) for c in nc if pd.notna(row[c])]
                except Exception:
                    continue
                if len(nums) != 6 or len(set(nums)) != 6 or not all(1 <= n <= 45 for n in nums):
                    continue
                try:
                    rnd = int(float(row[rc])) if rc is not None and pd.notna(row[rc]) else i+1
                except Exception:
                    rnd = i+1
                bonus = None
                if bc is not None and pd.notna(row[bc]):
                    try:
                        b = int(float(row[bc]))
                        if 1 <= b <= 45 and b not in nums:
                            bonus = b
                    except Exception:
                        pass
                rows.append(LottoRow(rnd, tuple(sorted(nums)), bonus))
            if len(rows) > len(best):
                best = rows
        if len(best) < 100:
            raise ValueError("유효한 당첨번호 회차를 충분히 인식하지 못했습니다.")
        return sorted({r.round_no:r for r in best}.values(), key=lambda x:x.round_no)

def zscores(counter):
    vals=[counter.get(n,0) for n in range(1,46)]
    m=sum(vals)/45
    sd=math.sqrt(sum((x-m)**2 for x in vals)/45) or 1
    return {n:(counter.get(n,0)-m)/sd for n in range(1,46)}

def gap_scores(history):
    last={n:-1 for n in range(1,46)}
    for i,d in enumerate(history):
        for n in d.numbers:last[n]=i
    gaps={n:len(history)-1-last[n] if last[n]>=0 else len(history) for n in range(1,46)}
    m=sum(gaps.values())/45
    sd=math.sqrt(sum((g-m)**2 for g in gaps.values())/45) or 1
    return {n:(gaps[n]-m)/sd for n in range(1,46)}

def number_scores(history, engine):
    weights={
        "baseline":(.45,.30,.15,.10),
        "survivor":(.30,.30,.20,.20),
        "consensus":(.25,.35,.30,.10),
    }[engine]
    allc=Counter(n for d in history for n in d.numbers)
    c240=Counter(n for d in history[-240:] for n in d.numbers)
    c100=Counter(n for d in history[-100:] for n in d.numbers)
    za,z240,z100,g=zscores(allc),zscores(c240),zscores(c100),gap_scores(history)
    return {n:weights[0]*za[n]+weights[1]*z240[n]+weights[2]*z100[n]+weights[3]*g[n] for n in range(1,46)}

def pair_scores(history):
    c=Counter()
    for d in history[-240:]: c.update(combinations(d.numbers,2))
    vals=list(c.values()) or [0]
    m=sum(vals)/len(vals)
    sd=math.sqrt(sum((v-m)**2 for v in vals)/len(vals)) or 1
    return {p:(v-m)/sd for p,v in c.items()}

def structure_score(combo):
    total=sum(combo); odd=sum(n%2 for n in combo); low=sum(n<=22 for n in combo)
    consecutive=sum(1 for a,b in zip(combo,combo[1:]) if b==a+1)
    same_end=6-len({n%10 for n in combo})
    return -abs(total-138)/25 -abs(odd-3)*.55 -abs(low-3)*.45 -max(0,consecutive-2)*1.2 -max(0,same_end-2)*.7

def generate_combinations(history, engine, count=100, candidate_size=18, stage=None):
    """상위 조합만 유지하는 메모리 절약형 생성기.

    stage(done, total, message)를 넘기면 긴 계산 중에도 UI에 진행 상황을 전달한다.
    주기적으로 아주 짧게 GIL을 양보하여 화면이 1%에서 멈춘 것처럼 보이지 않게 한다.
    """
    if stage:
        stage(0, 100, "번호 점수 계산")
    scores=number_scores(history,engine)
    ranked=sorted(range(1,46), key=lambda n:(-scores[n],n))
    candidates=ranked[:candidate_size]
    ps=pair_scores(history)

    total_combos=math.comb(len(candidates),6)
    keep=max(1200, min(5000, count*40))
    heap=[]
    for idx, combo in enumerate(combinations(candidates,6),1):
        ns=sum(scores[n] for n in combo)
        pair=sum(ps.get(tuple(sorted(p)),0) for p in combinations(combo,2))/15
        st=structure_score(combo)
        if engine=="baseline": total=ns+.35*pair+.8*st
        elif engine=="survivor": total=.85*ns+.25*pair+1.05*st
        else: total=.75*ns+.55*pair+.95*st

        item=(total, tuple(-n for n in combo), combo)
        if len(heap)<keep:
            heapq.heappush(heap,item)
        elif item>heap[0]:
            heapq.heapreplace(heap,item)

        if idx % 100 == 0:
            if stage:
                stage(idx,total_combos,f"조합 평가 {idx:,}/{total_combos:,}")
            time.sleep(0.002)

    pool=[(score,combo) for score,_,combo in heap]
    pool.sort(key=lambda x:(-x[0],x[1]))

    if stage:
        stage(95,100,"추천 조합 다양성 정리")
    selected=[]
    for _, combo in pool:
        cset=set(combo)
        if all(len(cset.intersection(prev)) <= 4 for prev in selected[-30:]):
            selected.append(combo)
            if len(selected) >= count:
                break
    if len(selected) < count:
        used=set(selected)
        for _, combo in pool:
            if combo not in used:
                selected.append(combo); used.add(combo)
                if len(selected) >= count:
                    break
    if stage:
        stage(100,100,"조합 생성 완료")
    return candidates,selected

def evaluate(draw,candidates,combos):
    actual=set(draw.numbers)
    ch=len(actual.intersection(candidates))
    vals=dict(first=0,second=0,third=0,fourth=0,fifth=0)
    maxm=0; win=0
    for combo in combos:
        m=len(actual.intersection(combo)); maxm=max(maxm,m)
        if m==6: vals["first"]+=1; win+=PAYOUT[6]
        elif m==5 and draw.bonus in combo: vals["second"]+=1; win+=PAYOUT["5b"]
        elif m==5: vals["third"]+=1; win+=PAYOUT[5]
        elif m==4: vals["fourth"]+=1; win+=PAYOUT[4]
        elif m==3: vals["fifth"]+=1; win+=PAYOUT[3]
    return RoundResult(draw.round_no,maxm,ch,int(ch==6),**vals,winnings=win,cost=len(combos)*PRICE_PER_LINE)

def run_engine(rows,engine,start,end,count,candidate_size,progress):
    by={r.round_no:r for r in rows}
    rounds=[r for r in range(start,end+1) if r in by]
    if not rounds: raise ValueError("검증 구간에 회차가 없습니다.")
    details=[]; digest=hashlib.sha256()
    for i,rnd in enumerate(rounds,1):
        history=[r for r in rows if r.round_no<rnd]
        progress(i-1,len(rounds),engine,rnd,"회차 준비")
        def stage(done,total,message):
            progress(i-1 + (done/max(total,1))*0.90, len(rounds), engine, rnd, message)
        cand,combos=generate_combinations(history,engine,count,candidate_size,stage)
        progress(i-0.05,len(rounds),engine,rnd,"당첨 결과 비교")
        rr=evaluate(by[rnd],cand,combos); details.append(rr)
        digest.update(f"{engine}|{rnd}|{cand}|{combos}|{rr.winnings}".encode())
        progress(i,len(rounds),engine,rnd,"회차 완료")
    cost=sum(x.cost for x in details); win=sum(x.winnings for x in details)
    s=EngineSummary(
        engine,len(details),(win-cost)/cost*100 if cost else 0,
        sum(x.first for x in details),sum(x.second for x in details),sum(x.third for x in details),
        sum(x.fourth for x in details),sum(x.fifth for x in details),
        sum(x.max_match for x in details)/len(details),
        sum(x.max_match>=4 for x in details),sum(x.max_match>=5 for x in details),
        sum(x.max_match>=6 for x in details),sum(x.candidate_all6 for x in details),
        sum(x.candidate_hits for x in details)/len(details),digest.hexdigest()
    )
    return s,details

class DashboardPage(QWidget):
    def __init__(self):
        super().__init__(); l=QVBoxLayout(self)
        t=QLabel("대시보드"); t.setObjectName("pageTitle"); l.addWidget(t)
        self.summary=QLabel("역대 로또 엑셀을 불러오세요."); self.summary.setObjectName("summaryCard"); self.summary.setAlignment(Qt.AlignCenter); self.summary.setMinimumHeight(180); l.addWidget(self.summary)
        l.addStretch()

class FrequencyPage(QWidget):
    def __init__(self):
        super().__init__(); l=QVBoxLayout(self)
        t=QLabel("번호 출현 빈도"); t.setObjectName("pageTitle"); l.addWidget(t)
        self.table=QTableWidget(45,3); self.table.setHorizontalHeaderLabels(["번호","출현 횟수","비율"]); self.table.verticalHeader().setVisible(False); l.addWidget(self.table)

class PairTriplePage(QWidget):
    def __init__(self):
        super().__init__(); l=QVBoxLayout(self)
        t=QLabel("페어·트리플 분석"); t.setObjectName("pageTitle"); l.addWidget(t)
        s=QSplitter(Qt.Horizontal)
        self.pair=QTableWidget(0,2); self.pair.setHorizontalHeaderLabels(["페어","횟수"])
        self.triple=QTableWidget(0,2); self.triple.setHorizontalHeaderLabels(["트리플","횟수"])
        s.addWidget(self.pair); s.addWidget(self.triple); l.addWidget(s)

class NumberInputPage(QWidget):
    def __init__(self):
        super().__init__(); l=QVBoxLayout(self)
        t=QLabel("번호 직접 입력"); t.setObjectName("pageTitle"); l.addWidget(t)
        self.box=QPlainTextEdit(); self.box.setPlaceholderText("예: 3 8 12 19 27 41"); l.addWidget(self.box)
        b=QPushButton("입력 번호 검사"); b.clicked.connect(self.check); l.addWidget(b)
        self.result=QLabel(""); self.result.setWordWrap(True); l.addWidget(self.result)
    def check(self):
        txt=self.box.toPlainText()
        for x in [",","\n","\t",";","/"]: txt=txt.replace(x," ")
        try: nums=[int(x) for x in txt.split()]
        except: self.result.setText("숫자만 입력하세요."); return
        bad=[n for n in nums if not 1<=n<=45]
        if bad: self.result.setText("범위 오류: "+", ".join(map(str,bad))); return
        self.result.setText(f"고유 번호 {len(set(nums))}개\n"+", ".join(map(str,sorted(set(nums)))))

class RecommendationPage(QWidget):
    def __init__(self):
        super().__init__(); self.rows=[]; l=QVBoxLayout(self)
        t=QLabel("추천 조합"); t.setObjectName("pageTitle"); l.addWidget(t)
        self.active_label=QLabel("현재 사용 엔진: baseline"); l.addWidget(self.active_label)
        self.engine=QComboBox(); self.engine.addItems(["baseline","survivor","consensus"]); l.addWidget(self.engine)
        b=QPushButton("현재 데이터로 TOP 20 생성"); b.clicked.connect(self.make); l.addWidget(b)
        self.table=QTableWidget(0,2); self.table.setHorizontalHeaderLabels(["순위","추천 조합"]); l.addWidget(self.table)
    def set_rows(self,rows): self.rows=rows
    def set_active_engine(self, engine):
        if engine not in ["baseline","survivor","consensus"]:
            return
        self.engine.setCurrentText(engine)
        self.active_label.setText(f"현재 사용 엔진: {engine}")
    def make(self):
        if not self.rows: QMessageBox.warning(self,"확인","먼저 엑셀을 불러오세요."); return
        _,combos=generate_combinations(self.rows,self.engine.currentText(),20,18)
        self.table.setRowCount(len(combos))
        for i,c in enumerate(combos):
            self.table.setItem(i,0,QTableWidgetItem(str(i+1)))
            self.table.setItem(i,1,QTableWidgetItem("  ".join(map(str,c))))

class AutoResearchPage(QWidget):
    engine_results = Signal(object)

    def __init__(self):
        super().__init__()
        self.rows=[]
        self.source=None
        self.process=None
        self.job_dir=None
        self.last_progress_time=0.0
        self.continuous_running=False
        self.stop_requested=False
        self.current_cycle=0
        self.total_cycles=1

        l=QVBoxLayout(self)
        t=QLabel("자동연구·검증센터"); t.setObjectName("pageTitle"); l.addWidget(t)
        f=QFormLayout()
        self.start=QSpinBox(); self.start.setRange(1,9999); self.start.setValue(1001)
        self.end=QSpinBox(); self.end.setRange(1,9999); self.end.setValue(1218)
        self.count=QSpinBox(); self.count.setRange(10,500); self.count.setValue(100)
        self.candidate=QSpinBox(); self.candidate.setRange(10,30); self.candidate.setValue(18)
        f.addRow("검증 시작",self.start)
        f.addRow("검증 종료",self.end)
        f.addRow("회차당 조합",self.count)
        f.addRow("후보 번호 수",self.candidate)

        self.continuous=QCheckBox("엔진 자동교체 후 다음 자동연구 계속 실행")
        self.cycles=QSpinBox(); self.cycles.setRange(1,10000); self.cycles.setValue(10)
        self.interval=QSpinBox(); self.interval.setRange(0,3600); self.interval.setValue(5)
        f.addRow("연속 자동연구",self.continuous)
        f.addRow("최대 반복 횟수",self.cycles)
        f.addRow("반복 대기(초)",self.interval)
        l.addLayout(f)

        buttons=QHBoxLayout()
        self.button=QPushButton("실제 자동연구 시작")
        self.button.setObjectName("primaryButton")
        self.button.clicked.connect(self.start_research)
        buttons.addWidget(self.button)

        self.stop_button=QPushButton("연속 자동연구 중지")
        self.stop_button.setEnabled(False)
        self.stop_button.clicked.connect(self.stop_research)
        buttons.addWidget(self.stop_button)
        l.addLayout(buttons)

        self.progress=QProgressBar()
        self.progress.setRange(0,100)
        l.addWidget(self.progress)

        self.status=QLabel("엑셀을 불러온 뒤 실행하세요.")
        l.addWidget(self.status)

        self.table=QTableWidget(0,8)
        self.table.setHorizontalHeaderLabels(["엔진","ROI","1등","2등","3등","4등","5등","평균최대일치"])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        l.addWidget(self.table)

        self.log=QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(180)
        l.addWidget(self.log)

        self.watchdog=QTimer(self)
        self.watchdog.setInterval(1000)
        self.watchdog.timeout.connect(self.watch_process)

    def set_data(self,rows,source):
        self.rows=rows
        self.source=source
        latest=max(r.round_no for r in rows)
        self.end.setValue(latest)
        self.start.setValue(max(1, latest-10))
        self.status.setText(f"{len(rows)}회 데이터 연결됨 · 최신 {latest}회")

    def start_research(self):
        if not self.rows:
            QMessageBox.warning(self,"확인","먼저 엑셀을 불러오세요.")
            return
        if self.process is not None or self.continuous_running:
            QMessageBox.information(self,"진행 중","자동연구가 이미 실행 중입니다.")
            return
        if self.start.value() > self.end.value():
            QMessageBox.warning(self,"범위 오류","검증 시작 회차가 종료 회차보다 큽니다.")
            return

        self.stop_requested=False
        self.current_cycle=0
        self.total_cycles=self.cycles.value() if self.continuous.isChecked() else 1
        self.continuous_running=True
        self.button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.run_cycle()

    def stop_research(self):
        self.stop_requested=True
        self.continuous_running=False
        self.stop_button.setEnabled(False)
        if self.process is not None:
            self.status.setText("중지 요청 중… 현재 연구 프로세스를 종료합니다.")
            self.process.kill()
        else:
            self.status.setText("연속 자동연구가 중지되었습니다.")
            self.button.setEnabled(True)

    def run_cycle(self):
        if self.stop_requested:
            self.finish_continuous("사용자 요청으로 연속 자동연구를 중지했습니다.")
            return
        if self.current_cycle >= self.total_cycles:
            self.finish_continuous(f"연속 자동연구 {self.total_cycles}회가 모두 완료되었습니다.")
            return

        self.current_cycle += 1

        available={r.round_no for r in self.rows}
        missing=[r for r in range(self.start.value(),self.end.value()+1) if r not in available]
        if missing:
            QMessageBox.warning(self,"데이터 확인",f"검증 구간에 없는 회차가 있습니다.\n첫 누락 회차: {missing[0]}")
            return

        self.progress.setValue(1)
        self.table.setRowCount(0)
        if self.current_cycle == 1:
            self.log.clear()
        self.log.appendPlainText(
            f"\n===== 연속 자동연구 {self.current_cycle}/{self.total_cycles} 시작 ====="
        )
        self.status.setText(
            f"자동연구 {self.current_cycle}/{self.total_cycles}회차 프로세스 시작 중…"
        )
        QApplication.processEvents()

        self.job_dir=Path(tempfile.mkdtemp(prefix="taegyeong_research_"))
        input_path=self.job_dir/"input.json"
        output_path=self.job_dir/"result.json"
        payload={
            "rows":[asdict(r) for r in self.rows],
            "start":self.start.value(),
            "end":self.end.value(),
            "count":self.count.value(),
            "candidate_size":self.candidate.value(),
            "cycle":self.current_cycle,
            "output_path":str(output_path),
        }
        input_path.write_text(json.dumps(payload,ensure_ascii=False),encoding="utf-8")

        self.process=QProcess(self)
        base_dir=runtime_base_dir()

        if getattr(sys,"frozen",False):
            worker_exe=base_dir/"research_worker.exe"
            if not worker_exe.exists():
                self.fail(f"연구 엔진 실행파일을 찾지 못했습니다.\n{worker_exe}")
                return
            self.process.setProgram(str(worker_exe))
            self.process.setArguments([str(input_path)])

            if sys.platform=="win32":
                try:
                    def hide_worker_console(args):
                        args.flags |= 0x08000000  # CREATE_NO_WINDOW
                    self.process.setCreateProcessArgumentsModifier(hide_worker_console)
                except Exception:
                    pass
        else:
            python_exe=Path(sys.executable)
            worker_script=Path(__file__).with_name("research_worker.py")
            self.process.setProgram(str(python_exe))
            self.process.setArguments(["-u",str(worker_script),str(input_path)])

        self.process.setWorkingDirectory(str(base_dir))
        env=QProcessEnvironment.systemEnvironment()
        env.insert("PYTHONUTF8","1")
        env.insert("PYTHONIOENCODING","utf-8")
        self.process.setProcessEnvironment(env)
        self.process.setProcessChannelMode(QProcess.SeparateChannels)
        self.process.readyReadStandardOutput.connect(self.read_stdout)
        self.process.readyReadStandardError.connect(self.read_stderr)
        self.process.errorOccurred.connect(self.process_error)
        self.process.finished.connect(self.process_finished)

        self.last_progress_time=time.time()
        self.watchdog.start()
        self.process.start()

        if not self.process.waitForStarted(5000):
            self.fail("자동연구 프로세스를 시작하지 못했습니다.\n"+self.process.errorString())
            return

        self.progress.setValue(2)
        self.status.setText(
            f"자동연구 {self.current_cycle}/{self.total_cycles}회차 시작 완료 · 계산 준비 중"
        )
        self.log.appendPlainText("연구 프로세스 PID: "+str(self.process.processId()))

    def read_stdout(self):
        if self.process is None:
            return
        raw=bytes(self.process.readAllStandardOutput()).decode("utf-8",errors="replace")
        for line in raw.splitlines():
            line=line.strip()
            if not line:
                continue
            try:
                event=json.loads(line)
            except Exception:
                self.log.appendPlainText(line)
                continue
            kind=event.get("type")
            if kind=="progress":
                pct=max(2,min(99,int(event.get("percent",2))))
                msg=str(event.get("message","연구 중"))
                self.progress.setValue(pct)
                self.status.setText(msg)
                self.last_progress_time=time.time()
            elif kind=="log":
                self.log.appendPlainText(str(event.get("message","")))
            elif kind=="error":
                self.log.appendPlainText(str(event.get("message","오류")))

    def read_stderr(self):
        if self.process is None:
            return
        msg=bytes(self.process.readAllStandardError()).decode("utf-8",errors="replace").strip()
        if msg:
            self.log.appendPlainText("[작업 오류]\n"+msg)
            self.last_progress_time=time.time()

    def process_error(self,_error):
        if self.process is not None:
            self.log.appendPlainText("프로세스 오류: "+self.process.errorString())

    def watch_process(self):
        if self.process is None:
            self.watchdog.stop()
            return
        elapsed=int(time.time()-self.last_progress_time)
        if elapsed >= 5:
            self.status.setText(f"계산 계속 진행 중 · 마지막 응답 {elapsed}초 전")
        if elapsed >= 120:
            self.log.appendPlainText("경고: 120초 동안 진행 신호가 없습니다. 작업을 종료합니다.")
            self.process.kill()

    def process_finished(self,exit_code,exit_status):
        self.watchdog.stop()
        proc=self.process
        self.process=None
        if proc is not None:
            raw=bytes(proc.readAllStandardOutput()).decode("utf-8",errors="replace")
            if raw.strip():
                for line in raw.splitlines():
                    try:
                        event=json.loads(line)
                        if event.get("type")=="progress":
                            self.progress.setValue(max(2,min(99,int(event.get("percent",2)))))
                            self.status.setText(str(event.get("message","연구 중")))
                    except Exception:
                        self.log.appendPlainText(line)
            err=bytes(proc.readAllStandardError()).decode("utf-8",errors="replace").strip()
            if err:
                self.log.appendPlainText("[종료 오류]\n"+err)

        result_path=self.job_dir/"result.json" if self.job_dir else None
        if exit_code != 0 or not result_path or not result_path.exists():
            if self.stop_requested:
                self.finish_continuous("사용자 요청으로 자동연구를 중지했습니다.")
            else:
                self.fail(f"자동연구가 완료되지 않았습니다. 종료코드: {exit_code}\n아래 로그를 확인하세요.")
            return

        try:
            payload=json.loads(result_path.read_text(encoding="utf-8"))
            summaries=[EngineSummary(**x) for x in payload["summaries"]]
            decisions=[tuple(x) for x in payload["decisions"]]
            details={
                engine:[RoundResult(**row) for row in rows]
                for engine,rows in payload["details"].items()
            }
            self.done(summaries,details,decisions)
        except Exception:
            self.fail("연구 결과를 읽는 중 오류가 발생했습니다.\n"+traceback.format_exc())

    def done(self,summaries,details,decisions):
        self.table.setRowCount(len(summaries))
        for i,s in enumerate(summaries):
            vals=[s.engine,f"{s.roi:.2f}%",s.first,s.second,s.third,s.fourth,s.fifth,f"{s.avg_max_match:.3f}"]
            for j,v in enumerate(vals):
                self.table.setItem(i,j,QTableWidgetItem(str(v)))

        out=Path(self.source).parent/"太炅_통합자동연구_결과"
        out.mkdir(exist_ok=True)
        payload={
            "app_version":APP_VERSION,
            "settings":{
                "start":self.start.value(),
                "end":self.end.value(),
                "count":self.count.value(),
                "candidate_size":self.candidate.value(),
                "cycle":self.current_cycle,
                "total_cycles":self.total_cycles
            },
            "summaries":[asdict(s) for s in summaries],
            "decisions":decisions
        }
        run_hash=hashlib.sha256(json.dumps(payload,ensure_ascii=False,sort_keys=True).encode()).hexdigest()[:16]
        rd=out/f"검증결과_주기{self.current_cycle:04d}_{run_hash}"
        rd.mkdir(exist_ok=True)
        (rd/"summary.json").write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding="utf-8")

        with open(rd/"summary.csv","w",newline="",encoding="utf-8-sig") as f:
            w=csv.DictWriter(f,fieldnames=list(asdict(summaries[0]).keys()))
            w.writeheader()
            for s in summaries:
                w.writerow(asdict(s))

        wb=Workbook()
        ws=wb.active
        ws.title="엔진요약"
        headers=list(asdict(summaries[0]).keys())
        ws.append(headers)
        for s in summaries:
            ws.append([asdict(s)[h] for h in headers])
        ds=wb.create_sheet("판정")
        ds.append(["엔진","우세지표수","고등수미저하","판정"])
        for d in decisions:
            ds.append(list(d))
        wb.save(rd/"검증결과.xlsx")

        self.log.appendPlainText("\n".join(f"{d[0]}: {d[3]} (우세 {d[1]}개)" for d in decisions))
        self.log.appendPlainText(f"\n저장 위치:\n{rd}")
        self.status.setText(
            f"자동연구 {self.current_cycle}/{self.total_cycles}회차 검증 완료"
        )
        self.progress.setValue(100)

        # 신호는 즉시 전달되므로 AI 엔진관리의 자동교체가 먼저 실행된다.
        self.engine_results.emit(summaries)
        self.log.appendPlainText(
            f"연구 {self.current_cycle}회 완료 · 엔진 자동교체 판정 전달 완료"
        )

        if (
            self.continuous.isChecked()
            and not self.stop_requested
            and self.current_cycle < self.total_cycles
        ):
            wait_ms=self.interval.value()*1000
            self.status.setText(
                f"{self.interval.value()}초 후 다음 자동연구 시작 "
                f"({self.current_cycle+1}/{self.total_cycles})"
            )
            QTimer.singleShot(wait_ms,self.run_cycle)
        else:
            self.finish_continuous(
                f"자동연구 {self.current_cycle}회 완료\n\n마지막 저장 위치:\n{rd}",
                show_message=True
            )

    def finish_continuous(self,message,show_message=False):
        self.continuous_running=False
        self.process=None
        self.button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.status.setText(message.splitlines()[0])
        self.log.appendPlainText(message)
        if show_message:
            QMessageBox.information(self,"완료",message)

    def fail(self,msg):
        self.watchdog.stop()
        self.progress.setValue(0)
        self.log.appendPlainText(msg)
        self.status.setText("오류 발생 · 아래 로그를 확인하세요.")
        self.continuous_running=False
        self.button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.process=None
        QMessageBox.critical(self,"자동연구 오류",msg[-1500:])


class EngineManagerPage(QWidget):
    active_engine_changed = Signal(str)
    def __init__(self):
        super().__init__()
        self.summaries=[]
        self.config_path=Path(__file__).with_name("engine_config.json")
        self.active_engine="baseline"
        self.auto_mode=False
        l=QVBoxLayout(self)
        t=QLabel("AI 엔진관리"); t.setObjectName("pageTitle"); l.addWidget(t)
        self.current=QLabel("현재 챔피언 엔진: baseline"); self.current.setObjectName("summaryCard"); l.addWidget(self.current)
        self.auto=QCheckBox("검증 완료 후 가장 좋은 엔진으로 자동 교체")
        self.auto.stateChanged.connect(self.auto_changed); l.addWidget(self.auto)
        self.table=QTableWidget(0,7)
        self.table.setHorizontalHeaderLabels(["엔진","종합점수","ROI","4+회차","5+회차","후보6포함","판정"])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        l.addWidget(self.table)
        row=QHBoxLayout()
        self.choose=QComboBox(); self.choose.addItems(["baseline","survivor","consensus"]); row.addWidget(self.choose)
        b=QPushButton("선택한 엔진을 챔피언으로 적용"); b.setObjectName("primaryButton"); b.clicked.connect(self.manual_apply); row.addWidget(b)
        l.addLayout(row)
        self.note=QLabel("자동교체가 꺼져 있으면 연구 결과를 보고 직접 엔진을 선택할 수 있습니다.")
        self.note.setWordWrap(True); l.addWidget(self.note)
        self.load_config()

    def score(self,s):
        return s.avg_max_match*100 + s.max4_plus_rounds*8 + s.max5_plus_rounds*25 + s.candidate_all6_rounds*12 + s.roi*.02

    def receive_results(self,summaries):
        self.summaries=list(summaries)
        ranked=sorted(self.summaries,key=self.score,reverse=True)
        self.table.setRowCount(len(ranked))
        best_score=self.score(ranked[0]) if ranked else 0
        for i,x in enumerate(ranked):
            status="최우수 후보" if i==0 else "비교 후보"
            vals=[x.engine,f"{self.score(x):.2f}",f"{x.roi:.2f}%",x.max4_plus_rounds,x.max5_plus_rounds,x.candidate_all6_rounds,status]
            for j,v in enumerate(vals): self.table.setItem(i,j,QTableWidgetItem(str(v)))
        if self.auto_mode and ranked:
            # 자동교체 보호: 현재 엔진보다 실제 종합점수가 높을 때만 바꾼다.
            current=next((x for x in ranked if x.engine==self.active_engine),None)
            if current is None or self.score(ranked[0]) > self.score(current):
                self.apply_engine(ranked[0].engine, "자동")

    def manual_apply(self):
        self.apply_engine(self.choose.currentText(), "수동")

    def auto_changed(self,state):
        self.auto_mode=bool(state)
        self.save_config()
        self.note.setText("자동교체 ON: 검증에서 현재 챔피언보다 좋은 엔진만 자동 적용합니다." if self.auto_mode
                          else "자동교체 OFF: 연구 결과를 보고 직접 엔진을 선택합니다.")

    def apply_engine(self,engine,mode):
        self.active_engine=engine
        self.current.setText(f"현재 챔피언 엔진: {engine} ({mode} 적용)")
        self.choose.setCurrentText(engine)
        self.save_config()
        self.active_engine_changed.emit(engine)

    def save_config(self):
        self.config_path.write_text(json.dumps({"active_engine":self.active_engine,"auto_mode":self.auto_mode},ensure_ascii=False,indent=2),encoding="utf-8")

    def load_config(self):
        try:
            d=json.loads(self.config_path.read_text(encoding="utf-8"))
            self.active_engine=d.get("active_engine","baseline")
            self.auto_mode=bool(d.get("auto_mode",False))
        except Exception:
            pass
        self.choose.setCurrentText(self.active_engine)
        self.auto.setChecked(self.auto_mode)
        self.current.setText(f"현재 챔피언 엔진: {self.active_engine}")

class PlaceholderPage(QWidget):
    def __init__(self,title,body):
        super().__init__(); l=QVBoxLayout(self)
        t=QLabel(title); t.setObjectName("pageTitle"); l.addWidget(t)
        b=QLabel(body); b.setObjectName("summaryCard"); b.setAlignment(Qt.AlignCenter); b.setWordWrap(True); b.setMinimumHeight(220); l.addWidget(b); l.addStretch()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__(); self.rows=[]; self.current_file=None
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}"); self.resize(1320,860)
        self.dashboard=DashboardPage(); self.freq=FrequencyPage(); self.pair=PairTriplePage(); self.input=NumberInputPage(); self.recommend=RecommendationPage(); self.research=AutoResearchPage(); self.engine_manager=EngineManagerPage()
        self.research.engine_results.connect(self.engine_manager.receive_results)
        self.engine_manager.active_engine_changed.connect(self.recommend.set_active_engine)
        self.recommend.set_active_engine(self.engine_manager.active_engine)
        self.stack=QStackedWidget()
        pages=[self.dashboard,self.freq,self.pair,PlaceholderPage("사진 분석","사진 OCR 모듈 연결 예정"),self.input,self.recommend,PlaceholderPage("조합 검사","역대 당첨조합 비교 기능"),self.research,self.engine_manager,PlaceholderPage("설정","가중치·고정수·제외수 설정")]
        for p in pages:self.stack.addWidget(p)
        root=QWidget(); lay=QHBoxLayout(root); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0); lay.addWidget(self.sidebar()); lay.addWidget(self.stack,1); self.setCentralWidget(root)
        self.menu(); self.theme(); self.statusBar().showMessage("준비됨")

    def sidebar(self):
        f=QFrame(); f.setObjectName("sidebar"); f.setFixedWidth(250); l=QVBoxLayout(f)
        logo=QLabel("太炅"); logo.setObjectName("logo"); logo.setAlignment(Qt.AlignCenter); l.addWidget(logo)
        sub=QLabel("Lotto Lab Ultimate"); sub.setAlignment(Qt.AlignCenter); l.addWidget(sub); l.addSpacing(16)
        names=["대시보드","번호 빈도","페어·트리플","사진 분석","번호 입력","추천 조합","조합 검사","자동연구","AI 엔진관리","설정"]
        for i,n in enumerate(names):
            b=QPushButton(n); b.clicked.connect(lambda checked=False,x=i:self.stack.setCurrentIndex(x)); l.addWidget(b)
        l.addStretch()
        b=QPushButton("엑셀 불러오기"); b.setObjectName("primaryButton"); b.clicked.connect(self.open_excel); l.addWidget(b)
        return f

    def menu(self):
        m=self.menuBar().addMenu("파일"); a=QAction("엑셀 불러오기",self); a.setShortcut("Ctrl+O"); a.triggered.connect(self.open_excel); m.addAction(a)
        h=self.menuBar().addMenu("도움말"); ab=QAction("프로그램 정보",self); ab.triggered.connect(lambda:QMessageBox.information(self,"정보",f"{APP_NAME}\n{APP_VERSION}\n추천과 실제 워크포워드 자동연구 통합판")); h.addAction(ab)

    def open_excel(self):
        fn,_=QFileDialog.getOpenFileName(self,"역대 로또 엑셀 선택","","Excel (*.xlsx *.xlsm *.xls)")
        if not fn:return
        try:
            self.rows=LottoDataParser.parse_excel(fn); self.current_file=Path(fn); self.update_analysis()
            latest=max(r.round_no for r in self.rows)
            self.dashboard.summary.setText(f"파일: {self.current_file.name}\n분석 회차: {len(self.rows)}개\n최신 회차: {latest}회\n추천·자동연구 데이터 연결 완료")
            self.recommend.set_rows(self.rows); self.research.set_data(self.rows,fn)
            self.statusBar().showMessage(f"{self.current_file.name} 분석 완료")
        except Exception as e:
            QMessageBox.critical(self,"오류",f"{e}\n\n{traceback.format_exc(limit=2)}")

    def update_analysis(self):
        total=len(self.rows); nc=Counter(); pc=Counter(); tc=Counter()
        for r in self.rows:
            nc.update(r.numbers); pc.update(combinations(r.numbers,2)); tc.update(combinations(r.numbers,3))
        ranked=sorted(range(1,46),key=lambda n:(-nc[n],n))
        for i,n in enumerate(ranked):
            self.freq.table.setItem(i,0,QTableWidgetItem(str(n))); self.freq.table.setItem(i,1,QTableWidgetItem(str(nc[n]))); self.freq.table.setItem(i,2,QTableWidgetItem(f"{nc[n]/total*100:.2f}%"))
        for table,data in [(self.pair.pair,pc.most_common(100)),(self.pair.triple,tc.most_common(100))]:
            table.setRowCount(len(data))
            for i,(c,v) in enumerate(data):
                table.setItem(i,0,QTableWidgetItem(" · ".join(map(str,c)))); table.setItem(i,1,QTableWidgetItem(str(v)))

    def theme(self):
        self.setStyleSheet("""
        QMainWindow,QWidget{background:#111;color:#F4F0E6;font-family:"Malgun Gothic";font-size:14px}
        #sidebar{background:#080808;border-right:1px solid #4A3A12}
        #logo{color:#D4AF37;font-size:44px;font-weight:800}
        #pageTitle{color:#D4AF37;font-size:27px;font-weight:800;padding:8px}
        #summaryCard{background:#1A1A1A;border:1px solid #4A3A12;border-radius:12px;padding:24px}
        QPushButton{background:#242424;color:#F4F0E6;border:1px solid #3A3A3A;border-radius:8px;padding:10px;text-align:left}
        QPushButton:hover{border-color:#D4AF37;background:#2E2817}
        #primaryButton{background:#D4AF37;color:#111;font-weight:800;text-align:center;border:none}
        QTableWidget,QPlainTextEdit,QSpinBox,QComboBox{background:#171717;color:#F4F0E6;border:1px solid #3C3C3C}
        QHeaderView::section{background:#2A2416;color:#F0D980;padding:8px}
        QProgressBar{background:#222;border:1px solid #444;border-radius:7px;text-align:center}
        QProgressBar::chunk{background:#D4AF37}
        """)

def main():
    append_runtime_log(
        "startup.log",
        f"START\nversion={APP_VERSION}\npython={sys.executable}\nfrozen={getattr(sys,'frozen',False)}"
    )
    app=QApplication(sys.argv)
    install_exception_logger()
    app.setFont(QFont("Malgun Gothic",10))
    w=MainWindow()
    w.show()
    append_runtime_log("startup.log","MAIN WINDOW SHOWN")
    code=app.exec()
    append_runtime_log("startup.log",f"EXIT code={code}")
    return code

if __name__=="__main__":
    raise SystemExit(main())
