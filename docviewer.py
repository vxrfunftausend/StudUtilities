import os
import sys
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from tkinter import *
import tkinter as tk
from typing import Dict
from tkinter import ttk
from pathlib import Path

TRANSLATIONS = {
    "ru": {
        "choose_action": "Выберите действие",
        "format_gost": "Подгон под ГОСТ",
        "replace_text": "Замена текста",
        "exit": "Выход",
        "back": "Назад",
        "refresh_list": "Обновить список",
        "apply_gost": "Подогнать под ГОСТ",
        "format_desc": "Просмотр и форматирование документов DOCX",
        "replace_desc": "Замена текста в поле ввода",
        "source_text": "Исходный текст:",
        "find_label": "Заменить:",
        "replace_with_label": "На:",
        "replace_btn": "Заменить",
        "copy_result": "Копировать результат",
        "clear_btn": "Очистить",
        "result_ro": "Результат (только для чтения):",
        "select_file_first": "Сначала выберите файл в списке.",
        "done_saved": "Готово.\nСохранено: ",
        "list_refreshed": "Список обновлён.",
        "save_failed": (
            "Не удалось сохранить '{name}'.\n\n"
            "Файл открыт в другой программе (например, Word) "
            "или защищён от записи.\n"
            "Закройте его и попробуйте снова."
        ),
        "process_failed": "Не удалось обработать '{name}':\n{err}",
        "file_unavailable": (
            "Файл недоступен:\n{path}\n\n"
            "Нажмите «Обновить список»."
        ),
        "empty_file": "Файл пустой",
        "file_not_found": "Файл не найден",
        "file_corrupt": "Файл повреждён или не является DOCX",
        "no_text": "Документ не содержит текста",
    },
    "en": {
        "choose_action": "Choose an action",
        "format_gost": "Format to State Standard",
        "replace_text": "Replace text",
        "exit": "Exit",
        "back": "Back",
        "refresh_list": "Refresh list",
        "apply_gost": "Apply State Standard formatting",
        "format_desc": "View and format DOCX documents",
        "replace_desc": "Replace text in the input field",
        "source_text": "Source text:",
        "find_label": "Find:",
        "replace_with_label": "Replace with:",
        "replace_btn": "Replace",
        "copy_result": "Copy result",
        "clear_btn": "Clear",
        "result_ro": "Result (read-only):",
        "select_file_first": "Please select a file from the list first.",
        "done_saved": "Done.\nSaved: ",
        "list_refreshed": "List refreshed.",
        "save_failed": (
            "Could not save '{name}'.\n\n"
            "The file is open in another program (e.g., Word) "
            "or is write-protected.\n"
            "Close it and try again."
        ),
        "process_failed": "Could not process '{name}':\n{err}",
        "file_unavailable": (
            "File unavailable:\n{path}\n\n"
            "Click 'Refresh list'."
        ),
        "empty_file": "File is empty",
        "file_not_found": "File not found",
        "file_corrupt": "File is corrupted or not a DOCX",
        "no_text": "Document contains no text",
    },
}

current_lang = "en"


def t(key: str) -> str:
    return TRANSLATIONS[current_lang].get(key, key)


_tr_widgets = []


def reg(widget, key: str):
    _tr_widgets.append((widget, key))
    widget.configure(text=t(key))
    return widget


def apply_language() -> None:
    for widget, key in _tr_widgets:
        try:
            widget.configure(text=t(key))
        except tk.TclError:
            pass


def toggle_language() -> None:
    global current_lang
    current_lang = "en" if current_lang == "ru" else "ru"
    apply_language()
    lang_btn.configure(text="EN" if current_lang == "ru" else "RU")


def read_docx_text(path: Path) -> str:
    size = path.stat().st_size
    if size == 0:
        raise ValueError(t("empty_file"))

    try:
        document = Document(str(path))
    except PackageNotFoundError:
        raise ValueError(t("file_not_found"))
    except Exception as e:
        raise ValueError(f"{t('file_corrupt')} ({type(e).__name__})")

    parts = [p.text for p in document.paragraphs if p.text.strip()]
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    parts.append(cell.text)

    text = "\n".join(parts)
    if not text.strip():
        raise ValueError(t("no_text"))

    return text


def iter_all_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def add_numbers(paragraph) -> None:
    run = paragraph.add_run()

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")

    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


def formatting(doc) -> None:
    for paragraph in iter_all_paragraphs(doc):
        paragraph.paragraph_format.line_spacing = 1.5
        paragraph.paragraph_format.first_line_indent = Mm(12.5)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for section in doc.sections:
        footer = section.footer
        try:
            footer.is_linked_to_previous = False
        except Exception:
            pass

        if footer.paragraphs:
            p = footer.paragraphs[0]
            for r in list(p.runs):
                r._element.getparent().remove(r._element)
        else:
            p = footer.add_paragraph()

        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_numbers(p)

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)
    for section in doc.sections:
        section.left_margin   = Mm(30)
        section.right_margin  = Mm(15)
        section.top_margin    = Mm(20)
        section.bottom_margin = Mm(20)


def filter_docx(files):
    result = []
    for f in files:
        if f.lower().endswith(".docx"):
            result.append(f)
    return result


def get_docx_map(folder: Path) -> Dict[str, Path]:
    if not folder.is_dir():
        return {}
    return {
        p.name: p for p in folder.iterdir()
        if p.is_file() and p.name.lower().endswith(".docx")
    }


def docx_to_text(path: Path) -> str:
    document = Document(str(path))
    return "\n".join(p.text for p in document.paragraphs)


DOCS_DIR = Path(__file__).resolve().parent / "docs"
files_map = get_docx_map(DOCS_DIR)


def show_text(content: str):
    text.configure(state="normal")
    text.delete("1.0", "end")
    text.insert("1.0", content)
    text.configure(state="disabled")


def apply_formatting() -> None:
    name = combobox.get()
    path = files_map.get(name)
    if path is None or not path.is_file():
        show_text(t("select_file_first"))
        return

    out_path = path.with_name(f"{path.stem}{path.suffix}")
    try:
        doc = Document(str(path))
        formatting(doc)
        doc.save(str(out_path))
    except PermissionError:
        show_text(t("save_failed").format(name=out_path.name))
        return
    except Exception as e:
        show_text(t("process_failed").format(
            name=name, err=f"{type(e).__name__}: {e}"
        ))
        return

    refresh_list()
    show_text(t("done_saved") + out_path.name)


def refresh_list():
    global files_map
    files_map = get_docx_map(DOCS_DIR)
    combobox["values"] = list(files_map.keys())
    combobox.set("")
    show_text(t("list_refreshed"))


def selected(event):
    name = combobox.get()
    path = files_map.get(name)
    if path is None or not path.is_file():
        show_text(t("file_unavailable").format(path=path))
        return
    try:
        content = read_docx_text(path)
    except ValueError as e:
        show_text(f"«{name}»\n\n{e}")
        return
    show_text(content)



root = Tk()
root.title("StudUtilities")
_base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
_ico = os.path.join(_base, "studutilities.ico")
if os.path.exists(_ico):
    root.iconbitmap(_ico)

topbar = ttk.Frame(root)
topbar.pack(fill="x", side="top")

lang_btn = ttk.Button(topbar, text="EN", width=6, command=toggle_language)
lang_btn.pack(side="right", padx=(0, 8), pady=(4, 0))

container = ttk.Frame(root)
container.pack(fill="both", expand=True)
container.rowconfigure(0, weight=1)
container.columnconfigure(0, weight=1)

main_frame = ttk.Frame(container)
header = ttk.Label(
    main_frame,
    text="StudUtilities",
    font=("Segoe UI", 18, "bold"),
    anchor="w",
    padding=(10, 8),
)
header.pack(fill="x")

reg(
    ttk.Label(main_frame, text="", foreground="#666"),
    "choose_action",
).pack(anchor="w", padx=10, pady=(0, 20))

menu_buttons = ttk.Frame(main_frame)
menu_buttons.pack(fill="x", padx=20, pady=10)

reg(
    ttk.Button(menu_buttons, text="",
               command=lambda: show_page(format_frame)),
    "format_gost",
).pack(fill="x", pady=6)

reg(
    ttk.Button(menu_buttons, text="",
               command=lambda: show_page(replace_frame)),
    "replace_text",
).pack(fill="x", pady=6)

reg(
    ttk.Button(menu_buttons, text="", command=root.destroy),
    "exit",
).pack(fill="x", pady=(20, 6))

format_frame = ttk.Frame(container)
format_header = ttk.Label(
    format_frame,
    text="",
    font=("Segoe UI", 16, "bold"),
    anchor="w",
    padding=(10, 8),
)
format_header.pack(fill="x")
reg(format_header, "format_gost")

reg(
    ttk.Label(format_frame, text="", foreground="#666"),
    "format_desc",
).pack(anchor="w", padx=10, pady=(0, 6))

files_map: Dict[str, Path] = get_docx_map(DOCS_DIR)

combobox = ttk.Combobox(format_frame, values=list(files_map.keys()))
combobox.pack(anchor=NW, fill=X, padx=6, pady=6)
combobox.bind("<<ComboboxSelected>>", selected)

toolbar = ttk.Frame(format_frame)
toolbar.pack(fill="x", padx=10, pady=(0, 6))

reg(ttk.Button(toolbar, text="",
               command=lambda: show_page(main_frame)),
    "back").pack(side="left")

reg(ttk.Button(toolbar, text="", command=refresh_list),
    "refresh_list").pack(side="left", padx=(6, 0))

reg(ttk.Button(toolbar, text="", command=apply_formatting),
    "apply_gost").pack(side="left", padx=(4, 0))

reg(ttk.Button(toolbar, text="", command=root.destroy),
    "exit").pack(side="right")

text_frame = ttk.Frame(format_frame)
text_frame.pack(fill="both", expand=True, padx=6, pady=0)

text = tk.Text(text_frame, wrap="word")
scroll = ttk.Scrollbar(text_frame, orient="vertical", command=text.yview)
text.configure(yscrollcommand=scroll.set)

text.pack(side="left", fill="both", expand=True)
scroll.pack(side="right", fill="y")

replace_frame = ttk.Frame(container)

replace_header = ttk.Label(
    replace_frame,
    text="",
    font=("Segoe UI", 16, "bold"),
    anchor="w",
    padding=(10, 8),
)
replace_header.pack(fill="x")
reg(replace_header, "replace_text")

reg(
    ttk.Label(replace_frame, text="", foreground="#666"),
    "replace_desc",
).pack(anchor="w", padx=10, pady=(0, 6))

replace_toolbar = ttk.Frame(replace_frame)
replace_toolbar.pack(fill="x", padx=10, pady=(0, 6))
reg(
    ttk.Button(replace_toolbar, text="",
               command=lambda: show_page(main_frame)),
    "back",
).pack(side="left")

replace_body = ttk.Frame(replace_frame)
replace_body.pack(fill="both", expand=True, padx=10, pady=(0, 10))

reg(ttk.Label(replace_body, text=""), "source_text").pack(anchor="w")

input_frame = ttk.Frame(replace_body)
input_frame.pack(fill="both", expand=True, pady=(4, 8))

input_text = tk.Text(input_frame, wrap="word", height=6)
input_scroll = ttk.Scrollbar(input_frame, orient="vertical",
                             command=input_text.yview)
input_text.configure(yscrollcommand=input_scroll.set)
input_text.pack(side="left", fill="both", expand=True)
input_scroll.pack(side="right", fill="y")

search_frame = ttk.Frame(replace_body)
search_frame.pack(fill="x", pady=(0, 8))
search_frame.columnconfigure(1, weight=1)
search_frame.columnconfigure(3, weight=1)

reg(ttk.Label(search_frame, text=""), "find_label").grid(
    row=0, column=0, sticky="w", padx=(0, 4)
)
find_entry = ttk.Entry(search_frame)
find_entry.grid(row=0, column=1, sticky="ew", padx=(0, 12))

reg(ttk.Label(search_frame, text=""), "replace_with_label").grid(
    row=0, column=2, sticky="w", padx=(0, 4)
)
replace_entry = ttk.Entry(search_frame)
replace_entry.grid(row=0, column=3, sticky="ew")


def do_replace():
    source = input_text.get("1.0", "end-1c")
    find = find_entry.get()
    repl = replace_entry.get()
    result = source if find == "" else source.replace(find, repl)
    output_text.configure(state="normal")
    output_text.delete("1.0", "end")
    output_text.insert("1.0", result)
    output_text.configure(state="disabled")


def copy_result():
    root.clipboard_clear()
    root.clipboard_append(output_text.get("1.0", "end-1c"))


def clear_replace():
    input_text.delete("1.0", "end")
    find_entry.delete(0, "end")
    replace_entry.delete(0, "end")
    output_text.configure(state="normal")
    output_text.delete("1.0", "end")
    output_text.configure(state="disabled")


actions_frame = ttk.Frame(replace_body)
actions_frame.pack(fill="x", pady=(0, 8))
reg(ttk.Button(actions_frame, text="", command=do_replace),
    "replace_btn").pack(side="left")
reg(ttk.Button(actions_frame, text="", command=copy_result),
    "copy_result").pack(side="left", padx=(6, 0))
reg(ttk.Button(actions_frame, text="", command=clear_replace),
    "clear_btn").pack(side="left", padx=(6, 0))

reg(ttk.Label(replace_body, text=""), "result_ro").pack(anchor="w")

output_frame = ttk.Frame(replace_body)
output_frame.pack(fill="both", expand=True, pady=(4, 0))

output_text = tk.Text(output_frame, wrap="word", height=6, state="disabled")
output_scroll = ttk.Scrollbar(output_frame, orient="vertical",
                              command=output_text.yview)
output_text.configure(yscrollcommand=output_scroll.set)
output_text.pack(side="left", fill="both", expand=True)
output_scroll.pack(side="right", fill="y")


if not os.path.isdir("docs"):
    os.makedirs("docs")

PAGE_SIZES = {
    main_frame: "400x440",
    format_frame: "700x900",
    replace_frame: "650x780",
}


def show_page(page: ttk.Frame):
    for child in container.winfo_children():
        child.pack_forget()
    page.pack(fill="both", expand=True)
    root.geometry(PAGE_SIZES.get(page, "700x900"))


show_page(main_frame)
root.mainloop()