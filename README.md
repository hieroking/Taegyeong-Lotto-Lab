# 太炅 Lotto Lab v30 — GitHub Actions Windows EXE

이 패키지를 GitHub 저장소의 **루트**에 업로드하면 Windows에서 자동으로 EXE를 빌드합니다.

## 자동 빌드
1. 저장소에서 **Actions**
2. **Windows EXE Build v30**
3. **Run workflow**
4. 완료 후 실행 화면 아래 **Artifacts**
5. `Taegyeong-Lotto-Lab-v30-Windows` 다운로드

압축 안에는 다음 두 실행파일이 들어갑니다.

- `Taegyeong_Lotto_Lab_v30.exe` — 사용자가 실행하는 프로그램
- `research_worker.exe` — 자동연구 전용 엔진

두 파일은 반드시 같은 폴더에 두어야 합니다.

## 검은 CMD 창
메인 프로그램은 PyInstaller `--windowed`로 빌드됩니다.  
자동연구 엔진은 `CREATE_NO_WINDOW`로 실행되므로 검은 CMD 창을 띄우지 않습니다.

## 오류 기록
프로그램과 같은 폴더에 다음 파일이 생성됩니다.

- `startup.log`
- `error.log`
