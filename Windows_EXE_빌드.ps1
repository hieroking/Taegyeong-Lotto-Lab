$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m py_compile app.py research_worker.py

pyinstaller --noconfirm --clean --onefile --console `
  --name research_worker `
  --collect-all pandas `
  --collect-all openpyxl `
  research_worker.py

pyinstaller --noconfirm --clean --onefile --windowed `
  --name "Taegyeong_Lotto_Lab_v30" `
  --collect-all PySide6 `
  --collect-all pandas `
  --collect-all openpyxl `
  app.py

New-Item -ItemType Directory -Force -Path release\Taegyeong_Lotto_Lab_v30 | Out-Null
Copy-Item dist\Taegyeong_Lotto_Lab_v30.exe release\Taegyeong_Lotto_Lab_v30\
Copy-Item dist\research_worker.exe release\Taegyeong_Lotto_Lab_v30\
Compress-Archive -Path release\Taegyeong_Lotto_Lab_v30 -DestinationPath release\Taegyeong_Lotto_Lab_v30_Windows.zip -Force

Write-Host "BUILD COMPLETE: release\Taegyeong_Lotto_Lab_v30_Windows.zip"
