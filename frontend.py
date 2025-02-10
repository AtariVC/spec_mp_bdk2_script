import sys
import os
import subprocess
import copy
import yaml
from pathlib import Path
from PyQt6.QtWidgets import QApplication, QWidget, QCheckBox, QVBoxLayout, QPushButton, QFileDialog, QLabel, QMessageBox, QHBoxLayout
from backend import load_data, prepare_data, merge_data, create_result_table, \
add_section_names, save_to_excel, filter_unwanted_sections, MK_creator, filter_unwanted_sections_MK,\
load_specification

class FileSelectionWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.spec_file = ''
        self.ekb_file = ''
        self.output_path = ''
        self.spec_path:str = ""
        self.ekb_path:str = ""
        self.init_ui()

    def save_to_config(self) -> None:
        settings_path: Path = Path(__file__).parent.joinpath("spec_path_config.yaml")
        config_data: dict[str, float | int | str] = {"spec_path": str(Path().joinpath(self.spec_path)), "ekb_path": str(Path().joinpath(self.ekb_path))}
        if not settings_path.exists():
                settings_path.touch()
        with open(str(settings_path), 'w', encoding='utf-8') as open_file:
            yaml.dump(config_data, open_file)
        open_file.close()

    def load_from_config(self) -> None:
        settings_path = Path(__file__).parent.joinpath("spec_path_config.yaml")
        if not settings_path.exists():
            QMessageBox.warning(None, "Ошибка", "Файл конфигурации не найден.")
            return
        # Чтение данных из YAML-файла
        with open(settings_path, "r", encoding="utf-8") as file:
            loaded_config_data = yaml.safe_load(file)

        self.spec_path = loaded_config_data.get("spec_path", "")
        self.ekb_path = loaded_config_data.get("ekb_path", "")
        file.close()

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
        self.process_button.setEnabled(True)  # Изначально кнопка неактивна

        # Чекбоксы "МП" и "МК"
        self.mp_checkbox = QCheckBox("МП")
        self.mp_checkbox.setChecked(True)
        self.mk_checkbox = QCheckBox("МК")
        self.mk_checkbox.setChecked(True)

        # Подключаем чекбоксы к функции обновления состояния кнопки
        self.mp_checkbox.stateChanged.connect(self.update_process_button_state)
        self.mk_checkbox.stateChanged.connect(self.update_process_button_state)

        # Метки для отображения путей
        self.spec_label = QLabel("Спецификация: Не выбрано")
        self.ekb_label = QLabel("ЭКБ: Не выбрано")

        # Компоновка: кнопки + чекбоксы в одной строке
        spec_layout = QHBoxLayout()
        spec_layout.addWidget(self.select_spec_button)
        spec_layout.addWidget(self.mp_checkbox)

        ekb_layout = QHBoxLayout()
        ekb_layout.addWidget(self.select_ekb_button)
        ekb_layout.addWidget(self.mk_checkbox)

        # Основной layout
        layout = QVBoxLayout()
        layout.addLayout(spec_layout)
        layout.addWidget(self.spec_label)
        layout.addLayout(ekb_layout)
        layout.addWidget(self.ekb_label)
        layout.addWidget(self.process_button)

        self.setLayout(layout)
        self.load_from_config()

    def update_process_button_state(self):
        """Обновление состояния кнопки в зависимости от чекбоксов."""
        if self.mp_checkbox.isChecked() or self.mk_checkbox.isChecked():
            self.process_button.setEnabled(True)
        else:
            self.process_button.setEnabled(False)

    def select_spec_file(self):
        """Выбор файла спецификации."""
        file_dialog = QFileDialog(self)
        self.spec_file, _ = file_dialog.getOpenFileName(self, "Выберите файл спецификации", str(Path().joinpath(self.spec_path)), "Excel Files (*.xlsx)")
        if self.spec_file:
            self.spec_label.setText(f"Спецификация: {self.spec_file}")

    def select_ekb_file(self):
        """Выбор файла ЭКБ."""
        file_dialog = QFileDialog(self)
        self.ekb_file, _ = file_dialog.getOpenFileName(self, "Выберите файл перечня ЭКБ", str(Path().joinpath(self.ekb_path)), "Excel Files (*.xlsx)")
        if self.ekb_file:
            self.ekb_label.setText(f"ЭКБ: {self.ekb_file}")

    def process_data(self):
        """Обработка данных и создание выходного файла."""
        if self.spec_file:
            self.spec_path = self.spec_file
            self.save_to_config()
        elif self.ekb_path:
            self.ekb_path = self.ekb_file
            self.save_to_config()
        if not self.spec_file:
            self.spec_label.setText("Пожалуйста, выберите файл спецификации")
        if not self.ekb_file:
            if self.mp_checkbox.isChecked():
                self.spec_label.setText("Пожалуйста, выберите файл перечня ЭКБ")
                return

        # Путь для сохранения выходного файла
        self.output_path = Path(self.spec_file).parent/"merged_output_MP.xlsx"
        MK_creator_path = Path(self.spec_file).parent/"grouped_book_MK.xlsx"

        if self.mp_checkbox.isChecked():
            specification, passports = load_data(self.spec_file, self.ekb_file)
            specification, passports = prepare_data(specification, passports)
            specification = filter_unwanted_sections(specification) 
            merged_data = merge_data(specification, passports)
            result = create_result_table(merged_data)
            final_data = add_section_names(result, specification)
            save_to_excel(final_data, self.output_path)
            # Отображаем путь к сохраненному файлу
            self.spec_label.setText(f"Файл сохранен: {self.output_path}")
            self.open_file(self.output_path)
        if self.mk_checkbox.isChecked():
            if not self.mp_checkbox.isChecked():
                specification = load_specification(self.spec_file)
            specification_MK = copy.deepcopy(specification)
            specification_MK = filter_unwanted_sections_MK(specification_MK)
            MK_creator(self.spec_file, specification_MK, MK_creator_path, specification)
            # Открытие выходного файла
            self.open_file(MK_creator_path)


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
