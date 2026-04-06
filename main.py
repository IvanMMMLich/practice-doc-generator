import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

sys.path.insert(0, os.path.dirname(__file__))

from excel_reader import load_students
from scripts.gen_izveshenie import generate as gen_izveshenie
from scripts.gen_otzyv import generate as gen_otzyv
from scripts.gen_zadanie import generate as gen_zadanie
from scripts.gen_tema import generate as gen_tema

DOCS = {
    "izveshenie": ("Извещение",    gen_izveshenie, "извещение.docx"),
    "otzyv":      ("Отзыв",        gen_otzyv,      "отзыв.docx"),
    "zadanie":    ("Инд. задание", gen_zadanie,    "индзадание.docx"),
    "tema":       ("Отчёт",        gen_tema,        "отчет.docx"),
}

BG      = "#f5f5f5"
CARD    = "#ffffff"
BORDER  = "#e0e0e0"
TEXT    = "#1a1a1a"
MUTED   = "#888888"
ACCENT  = "#2d6be4"
GREEN   = "#2e7d32"
GLITE   = "#e8f5e9"


def generate_for_student(student, base_dir, doc_keys):
    name = student.get(1, "Без_имени").replace(" ", "_")
    folder = os.path.join(base_dir, name)
    os.makedirs(folder, exist_ok=True)
    for key in doc_keys:
        _, fn, fname = DOCS[key]
        fn(student, os.path.join(folder, fname))
    return name


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Генератор документов практики")
        self.resizable(False, False)
        self.geometry("540x480")
        self.configure(bg=BG)

        self.excel_path = tk.StringVar()
        self.out_dir    = tk.StringVar()
        self.students   = []
        self.mode       = tk.StringVar(value="all")
        self.one_row    = tk.StringVar()
        self.range_from = tk.StringVar()
        self.range_to   = tk.StringVar()
        self.doc_vars   = {k: tk.BooleanVar(value=True) for k in DOCS}

        self._build()

    def _card(self, pady=(4, 4)):
        f = tk.Frame(self, bg=CARD,
                     highlightbackground=BORDER, highlightthickness=1)
        f.pack(fill="x", padx=16, pady=pady)
        return f

    def _label(self, parent, text, size=10, color=TEXT, bold=False):
        font = ("Helvetica", size, "bold" if bold else "normal")
        return tk.Label(parent, text=text, bg=CARD, fg=color, font=font)

    def _entry(self, parent, var, width=40, readonly=False):
        return tk.Entry(parent, textvariable=var, width=width,
                        state="readonly" if readonly else "normal",
                        bg="#f0f0f0" if readonly else CARD,
                        readonlybackground="#f0f0f0",
                        fg=TEXT, relief="flat",
                        font=("Helvetica", 10),
                        highlightbackground=BORDER,
                        highlightthickness=1)

    def _btn(self, parent, text, cmd, color=ACCENT, fg="#ffffff", width=None):
        b = tk.Button(parent, text=text, command=cmd,
                      bg=color, fg=fg,
                      activebackground=color, activeforeground=fg,
                      relief="flat", bd=0, cursor="hand2",
                      font=("Helvetica", 10, "bold"),
                      padx=10, pady=5)
        if width:
            b.config(width=width)
        return b

    def _log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _build(self):
        # шапка
        tk.Label(self, text="Генератор документов практики",
                 bg=BG, fg=TEXT,
                 font=("Helvetica", 13, "bold")).pack(pady=(14, 10))

        # шаг 1 — файл
        c1 = self._card()
        self._label(c1, "Шаг 1 — Excel файл", size=10,
                    color=MUTED).pack(anchor="w", padx=12, pady=(8, 2))
        r1 = tk.Frame(c1, bg=CARD)
        r1.pack(fill="x", padx=12, pady=(0, 4))
        self._entry(r1, self.excel_path, width=38,
                    readonly=True).pack(side="left", ipady=3)
        self._btn(r1, "Обзор...", self._pick_excel).pack(
            side="left", padx=(8, 0))

        self.status_lbl = tk.Label(c1, text="Загрузите файл для проверки",
                                   bg=CARD, fg=MUTED,
                                   font=("Helvetica", 10), anchor="w")
        self.status_lbl.pack(fill="x", padx=12, pady=(0, 8))

        # шаг 2 — студенты
        c2 = self._card()
        self._label(c2, "Шаг 2 — выбор студентов", size=10,
                    color=MUTED).pack(anchor="w", padx=12, pady=(8, 4))

        for val, lbl in [
            ("all",   "Все студенты"),
            ("one",   "Один студент (№ строки)"),
            ("range", "Диапазон строк (с … по …)"),
        ]:
            tk.Radiobutton(c2, text=lbl, variable=self.mode, value=val,
                           command=self._update_mode,
                           bg=CARD, fg=TEXT,
                           selectcolor="#e3f0ff",
                           activebackground=CARD,
                           font=("Helvetica", 10)).pack(anchor="w", padx=20)

        self.one_frame = tk.Frame(c2, bg=CARD)
        tk.Label(self.one_frame, text="Строка №",
                 bg=CARD, fg=MUTED,
                 font=("Helvetica", 10)).pack(side="left", padx=(24, 6))
        self._entry(self.one_frame, self.one_row,
                    width=6).pack(side="left", ipady=3)

        self.range_frame = tk.Frame(c2, bg=CARD)
        tk.Label(self.range_frame, text="С",
                 bg=CARD, fg=MUTED,
                 font=("Helvetica", 10)).pack(side="left", padx=(24, 4))
        self._entry(self.range_frame, self.range_from,
                    width=5).pack(side="left", ipady=3)
        tk.Label(self.range_frame, text="по",
                 bg=CARD, fg=MUTED,
                 font=("Helvetica", 10)).pack(side="left", padx=(8, 4))
        self._entry(self.range_frame, self.range_to,
                    width=5).pack(side="left", ipady=3)

        tk.Frame(c2, bg=CARD, height=8).pack()

        # шаг 3 — документы
        c3 = self._card()
        self._label(c3, "Шаг 3 — документы", size=10,
                    color=MUTED).pack(anchor="w", padx=12, pady=(8, 6))
        row3 = tk.Frame(c3, bg=CARD)
        row3.pack(fill="x", padx=12, pady=(0, 8))
        for key, (lbl, _, _) in DOCS.items():
            tk.Checkbutton(row3, text=lbl,
                           variable=self.doc_vars[key],
                           bg=CARD, fg=TEXT,
                           selectcolor="#e3f0ff",
                           activebackground=CARD,
                           font=("Helvetica", 10)).pack(side="left", padx=(0, 8))

        # шаг 4 — папка
        c4 = self._card()
        self._label(c4, "Шаг 4 — папка сохранения", size=10,
                    color=MUTED).pack(anchor="w", padx=12, pady=(8, 2))
        r4 = tk.Frame(c4, bg=CARD)
        r4.pack(fill="x", padx=12, pady=(0, 8))
        self._entry(r4, self.out_dir, width=38,
                    readonly=True).pack(side="left", ipady=3)
        self._btn(r4, "Изменить", self._pick_dir).pack(
            side="left", padx=(8, 0))

        # кнопка
        self._btn(self, "Сгенерировать документы",
                  self._run, color=GREEN).pack(
            fill="x", padx=16, pady=12, ipady=10)

        # лог
        tk.Label(self, text="Лог:", bg=BG, fg=MUTED,
                 font=("Helvetica", 9)).pack(anchor="w", padx=16)
        self.log_box = tk.Text(self, height=6, state="disabled",
                               font=("Courier", 10),
                               bg="#1e1e1e", fg="#d4d4d4",
                               relief="flat", bd=0,
                               highlightbackground=BORDER,
                               highlightthickness=1)
        self.log_box.pack(fill="x", padx=16, pady=(2, 14))

        self._update_mode()

    def _update_mode(self):
        self.one_frame.pack_forget()
        self.range_frame.pack_forget()
        if self.mode.get() == "one":
            self.one_frame.pack(anchor="w", pady=(4, 0))
        elif self.mode.get() == "range":
            self.range_frame.pack(anchor="w", pady=(4, 0))

    def _pick_excel(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel файлы", "*.xlsx *.xls")])
        if not path:
            return
        self.excel_path.set(path)
        if not self.out_dir.get():
            self.out_dir.set(os.path.dirname(path))
        try:
            self.students = load_students(path)
            self.status_lbl.config(
                text=f"Найдено студентов: {len(self.students)}",
                fg=GREEN)
        except Exception as e:
            self.status_lbl.config(text=f"Ошибка: {e}", fg="red")

    def _pick_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.out_dir.set(d)

    def _run(self):
        if not self.students:
            messagebox.showwarning("Нет данных", "Сначала загрузите Excel файл.")
            return
        out = self.out_dir.get()
        if not out:
            messagebox.showwarning("Нет папки", "Укажите папку для сохранения.")
            return
        doc_keys = [k for k, v in self.doc_vars.items() if v.get()]
        if not doc_keys:
            messagebox.showwarning("Нет документов", "Выберите хотя бы один документ.")
            return

        m = self.mode.get()
        if m == "all":
            targets = self.students
        elif m == "one":
            try:
                idx = int(self.one_row.get()) - 1
                targets = [self.students[idx]]
            except (ValueError, IndexError):
                messagebox.showerror("Ошибка", "Неверный номер строки.")
                return
        else:
            try:
                a = int(self.range_from.get()) - 1
                b = int(self.range_to.get())
                targets = self.students[a:b]
            except ValueError:
                messagebox.showerror("Ошибка", "Неверный диапазон.")
                return

        self._log(f"Генерация: {len(targets)} студент(ов), {len(doc_keys)} документ(ов)...")
        errors = 0
        for student in targets:
            try:
                name = generate_for_student(student, out, doc_keys)
                self._log(f"  OK  {name}")
            except Exception as e:
                self._log(f"  ERR {e}")
                errors += 1

        if errors == 0:
            self._log(f"Готово -> {out}")
            messagebox.showinfo("Готово", f"Сгенерировано: {len(targets)}\nПапка: {out}")
        else:
            self._log(f"Завершено с ошибками: {errors}")


if __name__ == "__main__":
    App().mainloop()