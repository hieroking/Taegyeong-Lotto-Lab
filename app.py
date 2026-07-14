import sys
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QLabel, QMainWindow, QWidget, QVBoxLayout

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("太炅 Lotto Lab Ultimate")
        self.resize(1000, 650)

        root = QWidget()
        layout = QVBoxLayout(root)

        logo = QLabel("太炅")
        logo.setAlignment(Qt.AlignCenter)
        logo.setStyleSheet("font-size:72px; font-weight:800; color:#D4AF37;")

        title = QLabel("Lotto Lab Ultimate")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:30px; font-weight:700; color:#F4F0E6;")

        note = QLabel("Windows 자동 빌드 기본 버전")
        note.setAlignment(Qt.AlignCenter)
        note.setStyleSheet("font-size:17px; color:#DDDDDD;")

        layout.addStretch()
        layout.addWidget(logo)
        layout.addWidget(title)
        layout.addWidget(note)
        layout.addStretch()

        root.setStyleSheet("background:#111111;")
        self.setCentralWidget(root)

app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec())
