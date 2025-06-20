import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QMessageBox,
    QTableWidget, QTableWidgetItem
)
from PyQt6.QtCore import Qt
from parser import parse_text_file
from converter import export_to_excel


class TextToExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Text2Excel Converter")
        self.setGeometry(100, 100, 600, 400)

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

        self.table = QTableWidget()
        self.layout.addWidget(self.table)

        self.setLayout(self.layout)

        self.input_path = None
        self.parsed_data = []

    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выбери текстовик", "", "Text Files (*.txt)")
        if path:
            self.input_path = path
            self.label.setText(f"Файл загружен: {path}")
            try:
                self.parsed_data = parse_text_file(path)
                self.populate_table(self.parsed_data)
                self.convert_button.setEnabled(bool(self.parsed_data))
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось прочитать файл:\n{e}")

    def populate_table(self, data):
        self.table.clear()
        if not data:
            return

        self.table.setRowCount(len(data))
        self.table.setColumnCount(len(data[0]))

        for row_idx, row in enumerate(data):
            for col_idx, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row_idx, col_idx, item)

    def convert_file(self):
        try:
            output_path, _ = QFileDialog.getSaveFileName(self, "Сохранить как", "", "Excel Files (*.xlsx)")
            if output_path:
                export_to_excel(self.parsed_data, output_path)
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
        QTableWidget {
            background-color: #2a2a2a;
            gridline-color: #444;
            border: 1px solid #444;
        }
        QHeaderView::section {
            background-color: #333;
            color: #ccc;
            padding: 4px;
            border: 1px solid #444;
        }
    """
    app.setStyleSheet(dark_stylesheet)

    window = TextToExcelApp()
    window.show()
    sys.exit(app.exec())
