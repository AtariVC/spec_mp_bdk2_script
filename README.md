____
# Скрипт для заполнения маршрутных паспортов и маршрутных карт 

## Описание приложения

Приложение для автоматической обработки и формирования таблиц с перечня ЭКБ для маршрутных паспортов (МП) и маршрутных карт (МК) из файлов спецификации и файла со списком компонентов и паспортов на эти ЭКБ. Основные функции:

1. Сравнение данных спецификации на плату и файла с перечнем ЭКБ
2. Формирования двух выходных файлов:
   - `output_MP.xlsx` – файл для заполнения перечня ЭКБ для МП
   - `output_MK.xlsx` – файл для заполнения перечня ЭКБ для МК
   - 
В полученных файлах будут таблицы уже подогнанные под формат перечней элементов для МП и МК. Каждая страница таблицы Excel соответствует новой страницы перечня ЭКБ в МП и МК, поэтому достаточно просто скопировать данные из Excel и вставить в соответствующий документ МП или МК.

В проекте будет представлен шаблон МП и МК. Перед заполнением перейти на страницу с перечнем ЭКБ и скопировать шаблон таблицы в свой МП или МК. После этого можно  копировать данные из `output_MP.xlsx` и `output_MK.xlsx`.

## Структура проекта

```
.
├──Шаблон МК и МП
	├── ЮМП_250_212_045_07_MK_Модуль_коммутатора.vsd  # Шаблон MK
	├── ЮМП_250_212_045_07_МП_Модуль_коммутатора.vsd  # Шаблон MП
├── .gitignore
├── backend.py               # Логика обработки данных
├── frontend.py              # Графический интерфейс
├── main.py                  # Точка запуска приложения
├── poetry.lock              # Файл блокировки зависимостей
├── pyproject.toml           # Конфигурация проекта и зависимости
└── README.md                # Документация
```
## Требования

- Python 3.8 или новее
- Установленные зависимости (указаны в pyproject.toml)

## Установка

1. Скопируйте репозиторий или загрузите файлы приложения
2. Установите зависимости с помощью Poetry ():

```bash
poetry install
```

## Запуск

1. Активируйте виртуальное окружение (если используете Poetry):

```bash
poetry shell
```

2. Запустите приложение:

```bash
python main.py
```

## Сборка в исполняемый файл (EXE)

Для создания standalone версии с помощью PyInstaller:

1. Убедитесь, что PyInstaller установлен:

```bash
poetry add pyinstaller --dev
```

2. Выполните сборку:

```bash
poetry run pyinstaller --onefile --windowed --icon=app.ico main.py
```

Собранный EXE будет в папке `dist`.

## Использование

1. Запустите приложение
2. Выберите файлы через интерфейс:
   - Спецификация (Excel)
   - Перечень ЭКБ (Excel)
3. Нажмите "Обработать данные"
4. Результаты автоматически откроются после обработки
