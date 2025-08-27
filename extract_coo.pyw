import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# pip install PyPDF2
from PyPDF2 import PdfReader


DATE_PAT = r"\b([A-Z]{3})\.?(\d{1,2}),(\d{4})\b"   # напр. AUG.19,2025 или AUG19,2025
MONTHS = {
    "JAN":"01","FEB":"02","MAR":"03","APR":"04","MAY":"05","JUN":"06",
    "JUL":"07","AUG":"08","SEP":"09","OCT":"10","NOV":"11","DEC":"12"
}

def convert_date(mmm_dd_yyyy: re.Match) -> str:
    mmm, dd, yyyy = mmm_dd_yyyy.group(1), mmm_dd_yyyy.group(2), mmm_dd_yyyy.group(3)
    return f"{dd.zfill(2)}.{MONTHS.get(mmm,'00')}.{yyyy[-2:]}"


def read_first_page_text(pdf_path: str) -> str:
    try:
        reader = PdfReader(pdf_path)
        page0 = reader.pages[0]
        text = page0.extract_text() or ""
        # Нормализуем пробелы/переносы
        text = " ".join(text.replace("\u00a0", " ").split())
        return text
    except Exception as e:
        return f""  # пусто => позже пометим ошибку


def find_certificate_no(text: str) -> str | None:
    m = re.search(r"Certificate\s*No\.?\s*([A-Z0-9][A-Z0-9/.\-]+)", text, flags=re.I)
    return m.group(1) if m else None


def find_certification_date(text: str) -> str | None:
    # Пытаемся взять дату именно из блока 12.Certification
    m_block = re.search(r"12\.Certification(.*?)(?:Copy|$)", text, flags=re.I)
    candidate = None
    if m_block:
        md = re.search(DATE_PAT, m_block.group(1))
        if md:
            candidate = convert_date(md)
    # Если не нашли в 12, берём последнюю дату в документе
    if not candidate:
        all_dates = list(re.finditer(DATE_PAT, text))
        if all_dates:
            candidate = convert_date(all_dates[-1])
    return candidate


def find_invoice_no(text: str) -> str | None:
    # Берём содержимое только раздела 10.Number and date of invoices
    m_blk = re.search(
        r"10\.Number\s*and\s*date\s*of\s*invoices(.*?)(?=11\.Declaration|12\.Certification|Copy|$)",
        text, flags=re.I
    )
    if not m_blk:
        return None
    block = m_blk.group(1)

    # Находим дату инвойса в блоке (обычно сразу после номера)
    md = re.search(DATE_PAT, block)
    # Поиск токенов до даты; если дату не нашли — берём весь блок
    pre = block[:md.start()] if md else block
    # Берём последние 100 символов перед датой, чтобы не поймать лишние названия товаров
    pre_tail = pre[-120:]

    # Ищем «похожий на номер инвойса» токен: буква+алфанум >= 6, часто после слеша
    candidates = re.findall(r"(?:/)?([A-Z][A-Z0-9]{5,})", pre_tail)
    return candidates[-1] if candidates else None


def process_files(paths: list[str]) -> list[str]:
    rows = []
    for p in paths:
        file_name = os.path.basename(p)
        text = read_first_page_text(p)
        if not text:
            rows.append(f"ERROR\tN/A\tN/A\t{file_name}")
            continue

        cert_no = find_certificate_no(text) or "N/A"
        cert_date = find_certification_date(text) or "N/A"
        invoice_no = find_invoice_no(text) or "N/A"

        rows.append(f"{cert_no}\t{cert_date}\t{invoice_no}\t{file_name}")
    return rows


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Извлечение данных из PDF (первая страница)")
        self.geometry("820x520")

        frm = ttk.Frame(self, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        self.btn_pick = ttk.Button(frm, text="Выбрать PDF…", command=self.pick_files)
        self.btn_pick.grid(row=0, column=0, sticky="w")

        self.btn_save = ttk.Button(frm, text="Обработать и сохранить…", command=self.run_and_save, state=tk.DISABLED)
        self.btn_save.grid(row=0, column=1, padx=8, sticky="w")

        self.lbl = ttk.Label(frm, text="Выбрано файлов: 0")
        self.lbl.grid(row=0, column=2, sticky="w", padx=8)

        self.txt = tk.Text(frm, height=22, wrap="none")
        self.txt.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(10,0))

        # Скроллы
        yscroll = ttk.Scrollbar(frm, orient="vertical", command=self.txt.yview)
        yscroll.grid(row=1, column=3, sticky="ns")
        self.txt.configure(yscrollcommand=yscroll.set)

        xscroll = ttk.Scrollbar(frm, orient="horizontal", command=self.txt.xview)
        xscroll.grid(row=2, column=0, columnspan=3, sticky="ew")
        self.txt.configure(xscrollcommand=xscroll.set)

        frm.columnconfigure(0, weight=0)
        frm.columnconfigure(1, weight=0)
        frm.columnconfigure(2, weight=1)
        frm.rowconfigure(1, weight=1)

        self.paths: list[str] = []

    def pick_files(self):
        paths = filedialog.askopenfilenames(
            title="Выберите PDF файлы",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not paths:
            return
        self.paths = list(paths)
        self.lbl.config(text=f"Выбрано файлов: {len(self.paths)}")
        self.btn_save.config(state=tk.NORMAL)
        self.txt.delete("1.0", tk.END)
        self.txt.insert(tk.END, "Файлы выбраны. Нажмите «Обработать и сохранить…»\n")

    def run_and_save(self):
        if not self.paths:
            messagebox.showinfo("Нет файлов", "Сначала выберите PDF-файлы.")
            return

        rows = process_files(self.paths)

        # Показать предпросмотр
        self.txt.delete("1.0", tk.END)
        for r in rows:
            self.txt.insert(tk.END, r + "\n")

        # Сохранить
        out_path = filedialog.asksaveasfilename(
            title="Сохранить результат",
            defaultextension=".txt",
            initialfile="result.txt",
            filetypes=[("Text file", "*.txt")]
        )
        if not out_path:
            return

        try:
            with open(out_path, "w", encoding="utf-8") as f:
                for r in rows:
                    f.write(r + "\n")
            messagebox.showinfo("Готово", f"Результат сохранён:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

if __name__ == "__main__":
    App().mainloop()
