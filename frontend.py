import sys
import os
import subprocess
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel
from backend import load_data, prepare_data, merge_data, create_result_table, add_section_names, save_to_excel, filter_unwanted_sections

class FileSelectionWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.spec_file = ''
        self.ekb_file = ''
        self.output_path = ''

        self.init_ui()

    def init_ui(self):
        """Инициализация интерфейса пользователя."""
        self.setWindowTitle("Выбор файлов для обработки")

        # Создание кнопок
        self.select_spec_button = QPushButton("Выбрать файл спецификации")
        self.select_spec_button.clicked.connect(self.select_spec_file)

        self.select_ekb_button = QPushButton("Выбрать файл ЭКБ")
        self.select_ekb_button.clicked.connect(self.select_ekb_file)

        self.process_button = QPushButton("Обработать данные")
        self.process_button.clicked.connect(self.process_data)

        # Метки для отображения путей
        self.spec_label = QLabel("Спецификация: Не выбрано")
        self.ekb_label = QLabel("ЭКБ: Не выбрано")

        # Размещение элементов
        layout = QVBoxLayout()
        layout.addWidget(self.select_spec_button)
        layout.addWidget(self.spec_label)
        layout.addWidget(self.select_ekb_button)
        layout.addWidget(self.ekb_label)
        layout.addWidget(self.process_button)

        self.setLayout(layout)

    def select_spec_file(self):
        """Выбор файла спецификации."""
        file_dialog = QFileDialog(self)
        self.spec_file, _ = file_dialog.getOpenFileName(self, "Выберите файл спецификации", "", "Excel Files (*.xlsx)")
        if self.spec_file:
            self.spec_label.setText(f"Спецификация: {self.spec_file}")

    def select_ekb_file(self):
        """Выбор файла ЭКБ."""
        file_dialog = QFileDialog(self)
        self.ekb_file, _ = file_dialog.getOpenFileName(self, "Выберите файл ЭКБ", "", "Excel Files (*.xlsx)")
        if self.ekb_file:
            self.ekb_label.setText(f"ЭКБ: {self.ekb_file}")

    def process_data(self):
        """Обработка данных и создание выходного файла."""
        if not self.spec_file or not self.ekb_file:
            self.spec_label.setText("Пожалуйста, выберите оба файла")
            return

        specification, passports = load_data(self.spec_file, self.ekb_file)
        specification, passports = prepare_data(specification, passports)
        specification = filter_unwanted_sections(specification)
        merged_data = merge_data(specification, passports)
        result = create_result_table(merged_data)
        final_data = add_section_names(result, specification)

        # Путь для сохранения выходного файла
        self.output_path = os.path.dirname(self.spec_file) + "/merged_output.xlsx"
        save_to_excel(final_data, self.output_path)
        
        # Отображаем путь к сохраненному файлу
        self.spec_label.setText(f"Файл сохранен: {self.output_path}")

        # Открытие выходного файла
        self.open_file(self.output_path)

    def open_file(self, file_path):
        """Открытие выходного файла в системе."""
        if sys.platform == 'win32':  # Для Windows
            os.startfile(file_path)
        elif sys.platform == 'darwin':  # Для macOS
            subprocess.run(['open', file_path])
        else:  # Для Linux
            subprocess.run(['xdg-open', file_path])

def main():
    app = QApplication(sys.argv)
    window = FileSelectionWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
