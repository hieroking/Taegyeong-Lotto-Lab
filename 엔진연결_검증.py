from pathlib import Path
import py_compile

app = Path(__file__).with_name("app.py")
py_compile.compile(str(app), doraise=True)
text = app.read_text(encoding="utf-8")
checks = {
    "버전 11.2.1": 'VERSION = "11.2.1-v27-v40-ocr-thread-fix"' in text,
    "v27 후보생성": "def v27_candidate_pool" in text,
    "자체추천 v27→v40": "def generate_self_v27_v40" in text,
    "v40 Elite Survival": "def select_diverse" in text and "v40 Elite Survival" in text,
    "자체추천 라우팅": "generate_self_v27_v40" in text,
    "공통 최종선별": "Recommender.select_diverse" in text,
    "OCR 백그라운드 스레드": "class OCRWorker(QThread)" in text,
    "OCR 비동기 시작": "self.ocr_worker.start()" in text,
    "OCR 중복 실행 차단": "현재 사진을 인식하고 있습니다" in text,
}
failed = [name for name, ok in checks.items() if not ok]
for name, ok in checks.items():
    print(f"[{'통과' if ok else '실패'}] {name}")
if failed:
    raise SystemExit("검증 실패: " + ", ".join(failed))
print("모든 정적 엔진 연결 검증을 통과했습니다.")
