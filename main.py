import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QLabel, QMessageBox,
    QTableWidget, QTableWidgetItem, QSpacerItem, QSizePolicy, QInputDialog
)
from PyQt6.QtCore import Qt
from parser import parse_text_file
from converter import export_to_excel


class TextToExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Text2Excel Converter")
        self.resize(700, 500)

        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(20, 20, 20, 20)
        self.layout.setSpacing(15)

        # Лейбл.
        self.label = QLabel("Выберите .txt файл для конвертации:")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.label)

        # Кнопки в горизонтальной обёртке.
        button_layout = QHBoxLayout()
        self.open_button = QPushButton("📂 Загрузить текстовик")
        self.open_button.clicked.connect(self.load_file)
        button_layout.addWidget(self.open_button)

        self.convert_button = QPushButton("📤 Конвертировать в Excel")
        self.convert_button.clicked.connect(self.convert_file)
        self.convert_button.setEnabled(False)
        button_layout.addWidget(self.convert_button)

        self.layout.addLayout(button_layout)

        # Таблица предпросмотра.
        self.table = QTableWidget()
        self.table.horizontalHeader().sectionDoubleClicked.connect(self.edit_header)
        self.table.setSizeAdjustPolicy(QTableWidget.SizeAdjustPolicy.AdjustToContents)
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


    def edit_header(self, index):
        old_header = self.table.horizontalHeaderItem(index).text()
        new_header, ok = QInputDialog.getText(self, "Переименовать столбец", f"Новое имя для столбца {index + 1}:", text=old_header)
        if ok and new_header.strip():
            self.table.setHorizontalHeaderItem(index, QTableWidgetItem(new_header.strip()))

    def populate_table(self, data):
        self.table.clear()
        if not data:
            return

        self.table.setRowCount(len(data))
        self.table.setColumnCount(len(data[0]))

        for col in range(len(data[0])):
            header_item = QTableWidgetItem(f"Столбец {col + 1}")
            self.table.setHorizontalHeaderItem(col, header_item)

        for row_idx, row in enumerate(data):
            for col_idx, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                self.table.setItem(row_idx, col_idx, item)

    def convert_file(self):
        try:
            output_path, _ = QFileDialog.getSaveFileName(self, "Сохранить как", "", "Excel Files (*.xlsx)")
            if output_path:
                headers = [
                    self.table.horizontalHeaderItem(col).text()
                    for col in range(self.table.columnCount())
                ]
                data = []
                for row in range(self.table.rowCount()):
                    data.append([
                        self.table.item(row, col).text() if self.table.item(row, col) else ''
                        for col in range(self.table.columnCount())
                    ])
                data.insert(0, headers) 
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
            padding: 8px 16px;
            border-radius: 6px;
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
            padding: 6px;
            border: 1px solid #444;
        }
    """
    app.setStyleSheet(dark_stylesheet)

    window = TextToExcelApp()
    window.show()
    sys.exit(app.exec())
