import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QMessageBox
)
from PyQt6.QtCore import Qt
from parser import parse_text_file
from converter import export_to_excel

class TextToExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Text2Excel Converter")
        self.setGeometry(100, 100, 400, 200)


        self.layout = QVBoxLayout()

        self.label = QLabel("Выберите .txt файл для конвертации:")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter) 
        self.layout.addWidget(self.label)

        self.open_button = QPushButton("📂 Загрузить текстовик")
        self.open_button.clicked.connect(self.load_file)
        self.layout.addWidget(self.open_button)

        self.convert_button = QPushButton("📤 Конвертировать в Excel")
        self.convert_button.clicked.connect(self.convert_file)
        self.convert_button.setEnabled(False)
        self.layout.addWidget(self.convert_button)

        self.setLayout(self.layout)

        self.input_path = None

    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выбери текстовик", "", "Text Files (*.txt)")
        if path:
            self.input_path = path
            self.label.setText(f"Файл загружен: {path}")
            self.convert_button.setEnabled(True)

    def convert_file(self):
        try:
            data = parse_text_file(self.input_path)

            output_path, _ = QFileDialog.getSaveFileName(self, "Сохранить как", "", "Excel Files (*.xlsx)")
            if output_path:
                export_to_excel(data, output_path)
                QMessageBox.information(self, "Успех", "Конвертация завершена!")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Что-то пошло не так: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    dark_stylesheet = """
        QWidget {
            background-color: #1e1e1e;
            color: #dddddd;
            font-family: Consolas;
            font-size: 14px;
        }
        QPushButton {
            background-color: #333;
            color: #fff;
            border: 1px solid #555;
            padding: 6px 12px;
            border-radius: 4px;
        }
        QPushButton:hover {
            background-color: #444;
        }
        QPushButton:pressed {
            background-color: #555;
        }
        QLabel {
            color: #aaa;
        }
    """ 
    app.setStyleSheet(dark_stylesheet)
    window = TextToExcelApp()
    window.show()
    sys.exit(app.exec())
