import os
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

def read_docx_text(path: Path) -> str:
    size = path.stat().st_size
    if size == 0:
        raise ValueError("Файл пустой")

    try:
        document = Document(str(path))
    except PackageNotFoundError:
        raise ValueError("Файл не найден")
    except Exception as e:
        raise ValueError(f"Файл повреждён или не является DOCX ({type(e).__name__})")

    parts = [p.text for p in document.paragraphs if p.text.strip()]
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    parts.append(cell.text)

    text = "\n".join(parts)
    if not text.strip():
        raise ValueError("Документ не содержит текста")

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

def filter(files):
    result =[]
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
        show_text("Сначала выберите файл в списке.")
        return

    out_path = path.with_name(f"{path.stem}{path.suffix}")
    try:
        doc = Document(str(path))
        formatting(doc)
        doc.save(str(out_path))
    except PermissionError:
        show_text(
            f"Не удалось сохранить '{out_path.name}'.\n\n"
            "Файл открыт в другой программе (например, Word) "
            "или защищён от записи.\n"
            "Закройте его и попробуйте снова."
        )
        return
    except Exception as e:
        show_text(f"Не удалось обработать '{name}':\n{type(e).__name__}: {e}")
        return

    refresh_list()
    show_text(f"Готово.\nСохранено: {out_path.name}")

def refresh_list():
    global files_map
    files_map = get_docx_map(DOCS_DIR)
    combobox["values"] = list(files_map.keys())
    combobox.set("")
    show_text("Список обновлён.")

def selected(event):
    name = combobox.get()
    path = files_map.get(name)
    if path is None or not path.is_file():
        show_text(f"Файл недоступен:\n{path}\n\nНажмите «Обновить список».")
        return
    try:
        text = read_docx_text(path)
    except ValueError as e:
        show_text(f"«{name}»\n\n{e}")
        return
    show_text(text)











root = Tk()
root.title("StudUtilities")
root.geometry("700x900")

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
ttk.Label(
    main_frame,
    text="Выберите действие",
    foreground="#666",
).pack(anchor="w", padx=10, pady=(0, 20))

menu_buttons = ttk.Frame(main_frame)
menu_buttons.pack(fill="x", padx=20, pady=10)







format_frame = ttk.Frame(container)
format_header = ttk.Label(
    format_frame,
    text="Подгон под ГОСТ",
    font=("Segoe UI", 16, "bold"),
    anchor="w",
    padding=(10, 8),
)
format_header.pack(fill="x")
ttk.Label(
    format_frame,
    text="Просмотр и форматирование документов DOCX",
    foreground="#666",
).pack(anchor="w", padx=10, pady=(0, 6))

files_map: Dict[str, Path] = get_docx_map(DOCS_DIR)

combobox = ttk.Combobox(format_frame, values=list(files_map.keys()))
combobox.pack(anchor=NW, fill=X, padx=6, pady=6)
combobox.bind("<<ComboboxSelected>>", selected)

toolbar = ttk.Frame(format_frame)
toolbar.pack(fill="x", padx=10, pady=(0, 6))

ttk.Button(toolbar, text="Назад", command=lambda: show_page(main_frame)).pack(side="left")
ttk.Button(toolbar, text="Обновить список", command=refresh_list).pack(side="left", padx=(6, 0))
ttk.Button(toolbar, text="Подогнать под ГОСТ", command=apply_formatting).pack(side="left", padx=(4, 0))
ttk.Button(toolbar, text="Выход", command=root.destroy).pack(side="right")

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
    text="Замена текста",
    font=("Segoe UI", 16, "bold"),
    anchor="w",
    padding=(10, 8),
)
replace_header.pack(fill="x")
ttk.Label(
    replace_frame,
    text="Раздел находится в разработке",
    foreground="#666",
).pack(anchor="w", padx=10, pady=(0, 6))

replace_body = ttk.Frame(replace_frame)
replace_body.pack(fill="both", expand=True, padx=20, pady=20)
ttk.Label(
    replace_body,
    text="WIP",
    justify="left",
).pack(anchor="w", pady=(0, 20))

ttk.Button(
    replace_body,
    text="Назад",
    command=lambda: show_page(main_frame),
).pack(anchor="w")


if not os.path.isdir("docs"):
    os.makedirs("docs")
dir_content = os.listdir('docs')
docs_list = filter(dir_content)


ttk.Button(
    menu_buttons,
    text="Подгон под ГОСТ",
    command=lambda: show_page(format_frame),
).pack(fill="x", pady=6)

ttk.Button(
    menu_buttons,
    text="Замена текста",
    command=lambda: show_page(replace_frame),
).pack(fill="x", pady=6)

ttk.Button(
    menu_buttons,
    text="Выход",
    command=root.destroy,
).pack(fill="x", pady=(20, 6))


def show_page(page: ttk.Frame):
    for child in container.winfo_children():
        child.pack_forget()
    page.pack(fill="both", expand=True)

show_page(main_frame)
root.mainloop()