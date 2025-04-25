import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter, PdfMerger

# ===== PDF 轉向功能 =====
ROTATION_MAP = {
    "left": 90,
    "right": -90,
    "180": 180
}

def rotate_pdfs(rotation_type):
    files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    if not files:
        return
    for file_path in files:
        try:
            reader = PdfReader(file_path)
            writer = PdfWriter()
            for page in reader.pages:
                page.rotate(ROTATION_MAP[rotation_type])
                writer.add_page(page)
            base = os.path.splitext(file_path)[0]
            new_file = f"{base}_rotated_{rotation_type}.pdf"
            with open(new_file, "wb") as f:
                writer.write(f)
        except Exception as e:
            messagebox.showerror("錯誤", f"處理檔案 {file_path} 時發生錯誤：\n{str(e)}")
            return
    messagebox.showinfo("完成", "所有PDF已成功轉向並儲存。")

# ===== PDF 分頁分割（每頁一檔）功能 =====
def split_pdf_pages(input_path, output_folder):
    try:
        reader = PdfReader(input_path)
        total_pages = len(reader.pages)
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        for i in range(total_pages):
            writer = PdfWriter()
            writer.add_page(reader.pages[i])
            output_filename = os.path.join(output_folder, f"{base_name}_page_{i+1}.pdf")
            with open(output_filename, "wb") as out_pdf:
                writer.write(out_pdf)
        messagebox.showinfo("完成", f"已成功分割 {total_pages} 頁！\n儲存於：{output_folder}")
    except Exception as e:
        messagebox.showerror("錯誤", str(e))

# ===== PDF 匯出特定頁功能 =====
def export_selected_pages(input_path, page_str):
    try:
        page_nums = [int(num.strip()) - 1 for num in page_str.split(",")]
        reader = PdfReader(input_path)
        writer = PdfWriter()
        for i in page_nums:
            if 0 <= i < len(reader.pages):
                writer.add_page(reader.pages[i])
            else:
                messagebox.showerror("錯誤", f"第 {i+1} 頁不存在。PDF 共有 {len(reader.pages)} 頁。")
                return
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                   filetypes=[("PDF files", "*.pdf")],
                                                   title="儲存為")
        if output_path:
            with open(output_path, "wb") as f:
                writer.write(f)
            messagebox.showinfo("完成", f"已成功匯出 {len(page_nums)} 頁。")
    except Exception as e:
        messagebox.showerror("錯誤", f"處理過程出現錯誤：\n{str(e)}")

# ===== PDF 合併功能 =====
def merge_pdf_files(files, save_path):
    try:
        pdf_merger = PdfMerger()
        for file in files:
            pdf_merger.append(file)
        pdf_merger.write(save_path)
        pdf_merger.close()
        messagebox.showinfo("完成", f"PDF 已成功合併至：\n{save_path}")
    except Exception as e:
        messagebox.showerror("錯誤", f"合併時發生錯誤：\n{e}")

# ===== GUI 主頁面設計 =====
def open_split_per_page():
    main_frame.pack_forget()
    split_page_frame.pack()

def open_rotate_page():
    main_frame.pack_forget()
    rotate_frame.pack()

def open_export_page():
    main_frame.pack_forget()
    export_page_frame.pack()

def open_merge_page():
    main_frame.pack_forget()
    merge_page_frame.pack()

def back_to_main(from_frame):
    from_frame.pack_forget()
    main_frame.pack()

# ===== 主畫面設計 =====
root = tk.Tk()
root.title("PDF 工具箱")
root.geometry("550x380")

main_frame = tk.Frame(root)
tk.Label(main_frame, text="歡迎使用 PDF 工具", font=("Arial", 16)).pack(pady=20)
tk.Button(main_frame, text="PDF 分頁分割器", width=30, command=open_split_per_page).pack(pady=5)
tk.Label(main_frame, text="PDF有幾頁就會分割成幾個檔案", font=("Arial", 10), fg="gray").pack()
tk.Button(main_frame, text="PDF 匯出指定頁", width=30, command=open_export_page).pack(pady=5)
tk.Button(main_frame, text="PDF 批次轉向工具", width=30, command=open_rotate_page).pack(pady=5)
tk.Button(main_frame, text="PDF 合併工具", width=30, command=open_merge_page).pack(pady=5)
tk.Label(main_frame, text="prototype by Yen Yu", font=("Arial", 10, "italic"), fg="gray").pack(side="bottom", pady=10)
main_frame.pack()

# ===== PDF 分頁分割器畫面 =====
split_page_frame = tk.Frame(root)
split_input = tk.Entry(split_page_frame, width=60)
split_output = tk.Entry(split_page_frame, width=60)

tk.Label(split_page_frame, text="選擇 PDF 檔案：").pack(pady=2)
split_input.pack()
tk.Button(split_page_frame, text="選擇檔案", command=lambda: split_input.insert(0, filedialog.askopenfilename())).pack()
tk.Label(split_page_frame, text="輸出資料夾：").pack(pady=2)
split_output.pack()
tk.Button(split_page_frame, text="選擇資料夾", command=lambda: split_output.insert(0, filedialog.askdirectory())).pack()
tk.Button(split_page_frame, text="開始分割", bg="green", fg="white",
          command=lambda: split_pdf_pages(split_input.get(), split_output.get())).pack(pady=5)
tk.Button(split_page_frame, text="返回主畫面", command=lambda: back_to_main(split_page_frame)).pack()

# ===== PDF 匯出特定頁畫面 =====
export_page_frame = tk.Frame(root)
export_file = tk.Entry(export_page_frame, width=60)
export_pages = tk.Entry(export_page_frame, width=30)

tk.Label(export_page_frame, text="PDF 檔案：").pack(pady=2)
export_file.pack()
tk.Button(export_page_frame, text="選擇檔案", command=lambda: export_file.insert(0, filedialog.askopenfilename())).pack()
tk.Label(export_page_frame, text="要匯出的頁碼（如：1,3,5）").pack(pady=2)
export_pages.pack()
tk.Button(export_page_frame, text="匯出頁面", command=lambda: export_selected_pages(export_file.get(), export_pages.get())).pack(pady=5)
tk.Button(export_page_frame, text="返回主畫面", command=lambda: back_to_main(export_page_frame)).pack()

# ===== PDF 轉向工具畫面 =====
rotate_frame = tk.Frame(root)
tk.Label(rotate_frame, text="選擇旋轉方向：", font=("Arial", 14)).pack(pady=10)
tk.Button(rotate_frame, text="向左旋轉 90°", width=20, command=lambda: rotate_pdfs("left")).pack(pady=5)
tk.Button(rotate_frame, text="向右旋轉 90°", width=20, command=lambda: rotate_pdfs("right")).pack(pady=5)
tk.Button(rotate_frame, text="旋轉 180°", width=20, command=lambda: rotate_pdfs("180")).pack(pady=5)
tk.Button(rotate_frame, text="返回主畫面", command=lambda: back_to_main(rotate_frame)).pack(pady=10)

# ===== PDF 合併工具畫面 =====
merge_page_frame = tk.Frame(root)
merge_listbox = tk.Listbox(merge_page_frame, width=60, height=10)
merge_listbox.pack()
tk.Button(merge_page_frame, text="選擇 PDF 檔案", command=lambda: [merge_listbox.delete(0, tk.END),
    [merge_listbox.insert(tk.END, f) for f in filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])]]).pack(pady=2)
tk.Button(merge_page_frame, text="合併 PDF", command=lambda: merge_pdf_files(
    merge_listbox.get(0, tk.END),
    filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
)).pack(pady=5)
tk.Button(merge_page_frame, text="返回主畫面", command=lambda: back_to_main(merge_page_frame)).pack()

root.mainloop()
