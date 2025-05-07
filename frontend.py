import sys
import os
import subprocess
import copy
import yaml
from pathlib import Path
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QFileDialog, QLabel, QMessageBox, QCheckBox)
from backend import load_data, prepare_data, merge_data, create_result_table, \
add_section_names, save_to_excel, filter_unwanted_sections, MK_creator, filter_unwanted_sections_MK

class FileSelectionWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.spec_file = ''
        self.ekb_file = ''
        self.output_path = ''
        self.spec_path: str = ""
        self.ekb_path: str = ""
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

        self.select_ekb_button = QPushButton("Выбрать файл перечня ЭКБ")
        self.select_ekb_button.clicked.connect(self.select_ekb_file)

        self.process_button = QPushButton("Обработать данные")
        self.process_button.clicked.connect(self.process_data)

        # Чекбоксы для выбора типа выходных файлов
        self.mp_checkbox = QCheckBox("output_MP")
        self.mp_checkbox.setChecked(True)  # По умолчанию выбран
        self.mk_checkbox = QCheckBox("output_MK")
        self.mk_checkbox.setChecked(True)  # По умолчанию выбран

        # Горизонтальный layout для чекбоксов и кнопки обработки
        process_layout = QHBoxLayout()
        process_layout.addWidget(self.mp_checkbox)
        process_layout.addWidget(self.mk_checkbox)
        process_layout.addStretch()  # Добавляем растягиваемое пространство
        process_layout.addWidget(self.process_button)

        # Метки для отображения путей
        self.spec_label = QLabel("Спецификация: Не выбрано!")
        self.ekb_label = QLabel("Перечень ЭКБ: Не выбрано!")

        # Размещение элементов
        layout = QVBoxLayout()
        layout.addWidget(self.select_spec_button)
        layout.addWidget(self.spec_label)
        layout.addWidget(self.select_ekb_button)
        layout.addWidget(self.ekb_label)
        layout.addLayout(process_layout)  # Добавляем горизонтальный layout вместо кнопки

        self.setLayout(layout)
        self.load_from_config()

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
        if self.spec_file and self.ekb_file:
            self.spec_path = self.spec_file
            self.ekb_path = self.ekb_file
            self.save_to_config()
        if not self.spec_file or not self.ekb_file:
            self.spec_label.setText("Пожалуйста, выберите оба файла")
            return
        
        try:
            specification, passports = load_data(self.spec_file, self.ekb_file)
            specification, passports = prepare_data(specification, passports)
        except Exception:
            self.spec_label.setText("Ошибка данных")
            self.ekb_label.setText("Ошибка данных")
        
        # Создаем папку output, если ее нет
        try:
            os.mkdir(Path(self.spec_file).parent/"output")
        except Exception:
            pass

        # Обработка для output_MP, если выбран соответствующий чекбокс
        if self.mp_checkbox.isChecked():
            specification_mp = copy.deepcopy(specification)
            specification_mp = filter_unwanted_sections(specification_mp)
            merged_data = merge_data(specification_mp, passports)
            result = create_result_table(merged_data)
            final_data = add_section_names(result, specification_mp)
            
            self.output_path = Path(self.spec_file).parent/"output/output_MP.xlsx"
            save_to_excel(final_data, self.output_path)
            self.spec_label.setText(f"Файл сохранен: {self.output_path}")
            self.open_file(self.output_path)

        # Обработка для output_MK, если выбран соответствующий чекбокс
        if self.mk_checkbox.isChecked():
            specification_mk = copy.deepcopy(specification)
            specification_mk = filter_unwanted_sections_MK(specification_mk)
            MK_creator_path = Path(self.spec_file).parent/"output/output_MK.xlsx"
            MK_creator(self.spec_file, specification_mk, MK_creator_path, specification)
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