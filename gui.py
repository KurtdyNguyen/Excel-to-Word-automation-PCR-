import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json

from pgd import process_pgd_excel

SETTINGS_FILE = "settings.json"

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"input_path": "", "output_folder": ""}

def save_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)

def select_input_file():
    path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if path:
        input_entry.delete(0, tk.END)
        input_entry.insert(0, path)

def select_output_folder():
    path = filedialog.askdirectory()
    if path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, path)

def run_processing():
    input_path = input_entry.get().strip()
    output_folder = output_entry.get().strip()

    if not os.path.isfile(input_path):
        messagebox.showerror("Lỗi", "Vui lòng chọn file Excel đầu vào hợp lệ.")
        return
    if not os.path.isdir(output_folder):
        messagebox.showerror("Lỗi", "Vui lòng chọn thư mục đầu ra hợp lệ.")
        return

    save_settings({"input_path": input_path, "output_folder": output_folder})

    try:
        results = process_pgd_excel(input_path, output_folder)
        if results:
            messagebox.showinfo("Thành công", f"Đã xuất {len(results)} file PGD vào:\n{output_folder}")
        else:
            messagebox.showwarning("Không có dữ liệu", "Không tìm thấy phôi để xuất.")
    except Exception as e:
        messagebox.showerror("Lỗi khi xử lý", str(e))

# === GUI SETUP ===
root = tk.Tk()
root.title("PGD Report Generator")
root.geometry("500x250")

settings = load_settings()

tk.Label(root, text="Chọn file Excel đầu vào:").pack(pady=5)
input_entry = tk.Entry(root, width=60)
input_entry.pack()
input_entry.insert(0, settings.get("input_path", ""))
tk.Button(root, text="Chọn File", command=select_input_file).pack(pady=5)

tk.Label(root, text="Chọn thư mục lưu báo cáo:").pack(pady=5)
output_entry = tk.Entry(root, width=60)
output_entry.pack()
output_entry.insert(0, settings.get("output_folder", ""))
tk.Button(root, text="Chọn Thư Mục", command=select_output_folder).pack(pady=5)

tk.Button(root, text="Tạo báo cáo PGD", command=run_processing, bg="green", fg="white", padx=20, pady=5).pack(pady=15)

root.mainloop()