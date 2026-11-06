# StudUtilities

[🇬🇧 English](#english) | [🇷🇺 Русский](#русский)

## English
**StudUtilities** is a Python program for working with Microsoft Word documents (.docx) that helps quickly bring documents into compliance with with Russian State Standards (GOST) requirements and conveniently replace characters in text.
### Features

- 📄 **State Standard formatting** — automatic DOCX formatting:
  - Font: Times New Roman, 14 pt
  - Line spacing: 1.5
  - First-line indent: 12.5 mm
  - Justified alignment
  - Margins: left 30 mm, right 15 mm, top/bottom 20 mm
  - Centered page numbers in the footer
  - Processes both body paragraphs and tables
- 🔁 **Text replacement** — fast substring replacement in source text with one-click copy of the result

### Requirements

- Python 3.8+
- Microsoft Word (or a compatible app) — not required, but useful to preview results
- OS: Windows (recommended), Linux, macOS

### Installation

1. Clone the repository:

```bash
git clone https://github.com/your-username/StudUtilities.git
cd StudUtilities
```

2. Install dependencies:

```bash
pip install python-docx
```

### Usage

1. Put the desired `.docx` files into the `docs` folder located next to the script (or `.exe`). If the folder doesn't exist, it will be created automatically.
2. Run the program:

```bash
python start.py
```

3. In the main menu, choose the action you need:

#### State Standard formatting
- Select a document from the dropdown list
- Preview its contents in the built-in text area
- Click **"Apply State Standard formatting"** — the file will be overwritten with the applied formatting
- The **"Refresh list"** button updates the list of files in the `docs` folder

#### Text replacement
- Paste or type the source text into the top field
- Specify what to find and what to replace it with
- Click **"Replace"** — the result appears in the field below
- The **"Copy result"** button copies the result to the clipboard

> ⚠️ **Note:** the program overwrites the original file. Make a backup if the original is important.


## Русский
**StudUtilities** — программа на Python для работы с документами Microsoft Word (`.docx`), которая помогает быстро привести документы в соответствие с требованиями ГОСТ и удобно заменять символы в тексте.
### Возможности

- 📄 **Подгон под ГОСТ** — автоматическое форматирование документов DOCX:
  - Шрифт: Times New Roman, 14 pt
  - Межстрочный интервал: 1.5
  - Абзацный отступ первой строки: 12.5 мм
  - Выравнивание по ширине
  - Поля: левое 30 мм, правое 15 мм, верхнее и нижнее по 20 мм
  - Нумерация страниц в нижнем колонтитуле (по центру)
  - Обрабатываются абзацы основного текста и таблиц
- 🔁 **Замена текста** — быстрая замена подстрок в исходном тексте с возможностью копировать результат в буфер обмена

### Требования

- Python 3.8+
- Microsoft Word (или совместимое приложение) — не требуется, но полезно для просмотра результатов
- ОС: Windows (рекомендуется), Linux, macOS

### Установка

1. Клонируйте репозиторий:

```bash
git clone https://github.com/your-username/StudUtilities.git
cd StudUtilities
```

2. Установите зависимости:

```bash
pip install python-docx
```

### Использование

1. Поместите нужные `.docx` файлы в папку `docs`, расположенную рядом со скриптом (или `.exe`). Если папки нет — она будет создана автоматически.
2. Запустите программу:

```bash
python start.py
```

3. В главном меню выберите нужное действие:

#### Подгон под ГОСТ
- Выберите документ из выпадающего списка
- Просмотрите содержимое во встроенном поле
- Нажмите **«Подогнать под ГОСТ»** — файл будет перезаписан с применённым форматированием
- Кнопка **«Обновить список»** обновляет перечень файлов в папке `docs`

#### Замена текста
- Вставьте или введите исходный текст в верхнее поле
- Укажите что заменить и на что
- Нажмите **«Заменить»** — результат появится в поле ниже
- Кнопка **«Копировать результат»** помещает результат в буфер обмена

> ⚠️ **Важно:** программа перезаписывает исходный файл. Сделайте резервную копию, если оригинал важен.
