import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_TITLE = "PINFL Test Data Generator"
TEST_PREFIX = "999"
MAX_SERIAL = 9_999_999_999

def test_check_digit(value: str) -> str:
    """Test-only checksum. NOT the official PINFL checksum algorithm."""
    return str(sum(int(ch) for ch in value) % 10)

def make_test_id(serial: int) -> str:
    body = TEST_PREFIX + f"{serial:010d}"  # 13 digits
    return body + test_check_digit(body)   # 14 digits

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("590x350")
        self.resizable(False, False)

        frm = ttk.Frame(self, padding=18)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Генератор фиктивных тестовых идентификаторов",
                  font=("Segoe UI", 14, "bold")).pack(anchor="w")
        ttk.Label(
            frm,
            text=("Создаёт 14-значные тестовые номера с зарезервированным префиксом 999. "
                  "Это НЕ ПИНФЛ и они не проходят официальную проверку."),
            wraplength=540
        ).pack(anchor="w", pady=(8, 16))

        grid = ttk.Frame(frm)
        grid.pack(fill="x")
        ttk.Label(grid, text="Начальный порядковый номер (0–9 999 999 999):").grid(row=0, column=0, sticky="w", pady=6)
        self.start_var = tk.StringVar(value="1")
        ttk.Entry(grid, textvariable=self.start_var, width=22).grid(row=0, column=1, sticky="e", padx=8)

        ttk.Label(grid, text="Количество записей:").grid(row=1, column=0, sticky="w", pady=6)
        self.count_var = tk.StringVar(value="1000")
        ttk.Entry(grid, textvariable=self.count_var, width=22).grid(row=1, column=1, sticky="e", padx=8)

        grid.columnconfigure(0, weight=1)

        self.preview_var = tk.StringVar(value="Пример: " + make_test_id(1))
        ttk.Label(frm, textvariable=self.preview_var, foreground="#245a86").pack(anchor="w", pady=(14, 6))

        note = ("Формат: 999 + 10 цифр порядкового номера + тестовая контрольная цифра. "
                "Контрольная цифра рассчитывается по сумме цифр по модулю 10; "
                "это специально НЕ официальный алгоритм ПИНФЛ.")
        ttk.Label(frm, text=note, wraplength=540).pack(anchor="w", pady=(0, 14))

        ttk.Button(frm, text="Создать XLSX", command=self.export).pack(anchor="w")
        ttk.Label(frm, text="В Excel каждая строка будет помечена TEST_ONLY / NOT_A_PINFL.",
                  font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(12, 0))

    def export(self):
        try:
            start = int(self.start_var.get().strip())
            count = int(self.count_var.get().strip())
            if start < 0 or start > MAX_SERIAL:
                raise ValueError("Начальный номер должен быть от 0 до 9 999 999 999.")
            if count < 1 or count > 1_000_000:
                raise ValueError("Количество должно быть от 1 до 1 000 000.")
            if start + count - 1 > MAX_SERIAL:
                raise ValueError("Диапазон выходит за пределы допустимого порядкового номера.")
        except ValueError as e:
            messagebox.showerror("Проверьте параметры", str(e))
            return

        path = filedialog.asksaveasfilename(
            title="Сохранить тестовые данные",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile="pinfl_test_data.xlsx"
        )
        if not path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "TEST DATA"
        ws.append(["TEST_IDENTIFIER", "STATUS", "SERIAL", "TEST_CHECK_DIGIT", "NOTICE"])
        for serial in range(start, start + count):
            identifier = make_test_id(serial)
            ws.append([identifier, "TEST_ONLY / NOT_A_PINFL", serial,
                       identifier[-1], "Synthetic ID; not a real PINFL"])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        header_fill = PatternFill("solid", fgColor="1F4E78")
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
        widths = [22, 30, 16, 20, 48]
        for i, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        # Store identifiers as text to preserve leading zeroes if format changes later.
        for cell in ws["A"][1:]:
            cell.number_format = "@"

        meta = wb.create_sheet("README")
        meta.append(["IMPORTANT"])
        meta.append(["These are synthetic test identifiers, not PINFLs."])
        meta.append(["Reserved test prefix: 999."])
        meta.append(["The final digit uses a test-only checksum (sum of digits modulo 10)."])
        meta.append(["The official PINFL checksum algorithm is intentionally not implemented."])
        meta.append(["Do not use these values to identify real people or in production systems."])
        meta.column_dimensions["A"].width = 105
        meta["A1"].font = Font(bold=True, color="C00000", size=14)
        wb.save(path)
        messagebox.showinfo("Готово", f"Создано записей: {count}\nФайл: {path}")

if __name__ == "__main__":
    App().mainloop()
