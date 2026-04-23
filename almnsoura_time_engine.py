import os
import sys
import threading
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog, Menu
from PIL import Image

# =========================
# App Config
# =========================
APP_NAME = "Almnsoura simple attendance"
APP_TAGLINE = "نظام إدارة الحضور والأسماء الذكي"
DEFAULT_OUT_SUFFIX = "_Odoo.xlsx"

MAIN_FONT = "Segoe UI"   # خط ويندوز افتراضي – دعم ممتاز للعربية

# =========================
# Paths
# =========================
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

NAME_MAP_FILE = os.path.join(BASE_DIR, "name_map.xlsx")

# =========================
# UI Theme
# =========================
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

BRAND = "#2563EB"
BG = "#F8FAFC"
CARD = "#FFFFFF"
BORDER = "#E2E8F0"
TEXT_MUTED = "#64748B"
DANGER = "#EF4444"
DANGER_LIGHT = "#FEF2F2"

# =========================
# Right Click Menu
# =========================
def show_right_click_menu(event):
    entry = event.widget
    menu = Menu(None, tearoff=0)
    menu.add_command(label="نسخ", command=lambda: entry.event_generate("<<Copy>>"))
    menu.add_command(label="لصق", command=lambda: entry.event_generate("<<Paste>>"))
    menu.add_command(label="قص", command=lambda: entry.event_generate("<<Cut>>"))
    menu.add_separator()
    menu.add_command(label="تحديد الكل", command=lambda: entry.select_range(0, 'end'))
    menu.tk_popup(event.x_root, event.y_root)

# =========================
# Data Helpers
# =========================
def load_raw_map():
    if os.path.exists(NAME_MAP_FILE):
        try:
            df = pd.read_excel(NAME_MAP_FILE)
            df.columns = [c.strip().lower() for c in df.columns]
            return df[["source_name", "odoo_name", "branch"]]
        except:
            pass
    return pd.DataFrame(columns=["source_name", "odoo_name", "branch"])

def save_raw_map(df):
    try:
        df.to_excel(NAME_MAP_FILE, index=False)
        return True
    except:
        messagebox.showerror("خطأ", "يرجى إغلاق ملف name_map.xlsx أولاً")
        return False

# =========================
# Employee Dialog
# =========================
class EmployeeDialog(ctk.CTkToplevel):
    def __init__(self, parent, branches, callback, initial=None):
        super().__init__(parent)
        self.title("بيانات الموظف")
        self.geometry("400x500")
        self.configure(fg_color=BG)
        self.callback = callback

        box = ctk.CTkFrame(self, fg_color="transparent")
        box.pack(fill="both", expand=True, padx=30, pady=30)

        ctk.CTkLabel(box, text="بيانات الموظف", font=(MAIN_FONT, 18, "bold")).pack(pady=20)

        ctk.CTkLabel(box, text="الاسم في جهاز البصمة (إنجليزي)").pack(anchor="e")
        self.src = ctk.CTkEntry(box, height=40, justify="right")
        self.src.pack(fill="x", pady=10)
        self.src.bind("<Button-3>", show_right_click_menu)

        ctk.CTkLabel(box, text="الاسم في أودو (عربي)").pack(anchor="e")
        self.dst = ctk.CTkEntry(box, height=40, justify="right")
        self.dst.pack(fill="x", pady=10)
        self.dst.bind("<Button-3>", show_right_click_menu)

        ctk.CTkLabel(box, text="الفرع").pack(anchor="e")
        self.branch = ctk.CTkOptionMenu(box, values=branches)
        self.branch.pack(fill="x", pady=15)

        if initial:
            self.src.insert(0, initial["source_name"])
            self.dst.insert(0, initial["odoo_name"])
            self.branch.set(initial["branch"])

        ctk.CTkButton(
            box,
            text="حفظ",
            height=45,
            fg_color="#10B981",
            command=self.save
        ).pack(fill="x", pady=20)

    def save(self):
        s = self.src.get().strip()
        d = self.dst.get().strip()
        b = self.branch.get()
        if not s or not d:
            messagebox.showwarning("تنبيه", "يرجى تعبئة جميع الحقول")
            return
        self.callback(s, d, b)
        self.destroy()

# =========================
# Main App
# =========================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("820x650")
        self.configure(fg_color=BG)

        self.input_path = ctk.StringVar()
        self.output_path = ctk.StringVar()

        self.build_ui()

    def build_ui(self):
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=40, pady=40)

        ctk.CTkLabel(
            main,
            text=APP_NAME,
            font=(MAIN_FONT, 24, "bold")
        ).pack(anchor="w")

        ctk.CTkLabel(
            main,
            text=APP_TAGLINE,
            font=(MAIN_FONT, 13),
            text_color=TEXT_MUTED
        ).pack(anchor="w", pady=(0, 30))

        self.field(main, "ملف تقرير الحضور", self.input_path, self.pick_input)
        self.field(main, "مسار حفظ الناتج", self.output_path, self.pick_output)

        self.btn = ctk.CTkButton(
            main,
            text="بدء التحويل الآن",
            height=55,
            fg_color=BRAND,
            font=(MAIN_FONT, 16, "bold"),
            command=self.start
        )
        self.btn.pack(fill="x", pady=30)

    def field(self, parent, label, var, cmd):
        ctk.CTkLabel(parent, text=label).pack(anchor="e")
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=10)

        e = ctk.CTkEntry(row, textvariable=var, justify="right")
        e.pack(side="left", fill="x", expand=True, padx=(0, 10))
        e.bind("<Button-3>", show_right_click_menu)

        ctk.CTkButton(row, text="اختيار", width=80, command=cmd).pack(side="right")

    def pick_input(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if p:
            self.input_path.set(p)
            self.output_path.set(p.replace(".xlsx", DEFAULT_OUT_SUFFIX))

    def pick_output(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if p:
            self.output_path.set(p)

    def start(self):
        if not self.input_path.get() or not self.output_path.get():
            messagebox.showwarning("تنبيه", "يرجى اختيار الملفات أولاً")
            return
        self.btn.configure(state="disabled", text="جاري المعالجة...")
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        try:
            map_df = load_raw_map()
            if map_df.empty:
                messagebox.showerror("خطأ", "قاعدة بيانات الموظفين فارغة")
                return

            name_map = {
                str(k).strip().lower(): str(v).strip()
                for k, v in zip(map_df["source_name"], map_df["odoo_name"])
            }

            raw = pd.read_excel(self.input_path.get()).astype(str)
            cols = {c.lower(): c for c in raw.columns}

            name_col = next((cols[c] for c in cols if "name" in c or "اسم" in c), None)
            date_col = next((cols[c] for c in cols if "date" in c or "تاريخ" in c), None)

            if not name_col or not date_col:
                messagebox.showerror("خطأ", "الأعمدة غير موجودة")
                return

            out = []
            for _, r in raw.iterrows():
                key = r[name_col].strip().lower()
                if key in name_map:
                    out.append({
                        "Employee": name_map[key],
                        "Date": r[date_col]
                    })

            if out:
                pd.DataFrame(out).to_excel(self.output_path.get(), index=False)
                messagebox.showinfo("نجاح", "تم التحويل بنجاح")
            else:
                messagebox.showwarning("تنبيه", "لا توجد أسماء مطابقة")

        finally:
            self.btn.configure(state="normal", text="بدء التحويل الآن")

# =========================
if __name__ == "__main__":
    App().mainloop()
