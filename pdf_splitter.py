import tkinter as tk
from tkinter import messagebox, filedialog
import subprocess
import sys
import os
import urllib.request
import pandas as pd
import shutil
import re
from PyPDF2 import PdfReader, PdfWriter
import zipfile

# 获取 Python 安装目录
PYTHON_INSTALL_DIR = os.path.dirname(sys.executable)
SCRIPTS_DIR = os.path.join(PYTHON_INSTALL_DIR, 'Scripts')

# 确保 Scripts 目录存在
os.makedirs(SCRIPTS_DIR, exist_ok=True)

PYTHON_PATH = os.path.join(PYTHON_INSTALL_DIR, "python.exe")

def get_documents_folder():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as key:
            return winreg.QueryValueEx(key, "Personal")[0]
    except:
        return os.path.join(os.path.expanduser("~"), "Documents")

OUTPUT_FOLDER = os.path.join(get_documents_folder(), "PDF_Splitter_Output")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def install_libraries():
    try:
        subprocess.run([PYTHON_PATH, "-m", "pip", "install", "pandas", "pypdf2", "openpyxl"], check=True)
        return True
    except subprocess.CalledProcessError:
        return False

def setup_environment():
    try:
        # 检查 Python 是否已安装
        subprocess.run([PYTHON_PATH, "--version"], check=True, capture_output=True)
    except FileNotFoundError:
        # Python 未安装，提示用户安装
        if messagebox.askyesno("Python 未安装", "是否安装 Python？"):
            python_installer = "python-3.12.0-amd64.exe"  # 确保这个文件在同一目录下
            subprocess.run([python_installer, "/quiet", "InstallAllUsers=1", "PrependPath=1", "Include_pip=1"])
        else:
            return False

    # 安装必要的库
    if install_libraries():
        messagebox.showinfo("成功", "必要的库已安装完成！")
        return True
    else:
        messagebox.showerror("错误", "安装库时出错")
        return False

def correct_pdf_name(file_path):
    file_name = os.path.basename(file_path)
    corrected_name = re.sub(r'(YY\d+-\d+-\d+-\d+).*\.pdf$', r'\1.pdf', file_name)
    return corrected_name

def process_excel():
    excel_path = filedialog.askopenfilename(title="请选择档案目录（excel文件）", filetypes=[("Excel files", "*.xls;*.xlsx")])
    if not excel_path:
        return

    df = pd.read_excel(excel_path)
    input_xlsx_path = os.path.join(SCRIPTS_DIR, 'input.xlsx')
    df.to_excel(input_xlsx_path, index=False)
    messagebox.showinfo("成功", f"Excel文件已保存为 {input_xlsx_path}")

    if messagebox.askyesno("确认", "是否需要选择要拆分的PDF文件？"):
        process_pdf()
    else:
        os.startfile(SCRIPTS_DIR)
        if messagebox.askyesno("确认", "请检查文件名是否正确。是否正确？"):
            if messagebox.askyesno("确认", "是否需要拆分文档？"):
                run_pdf_splitter()

def process_pdf():
    pdf_paths = filedialog.askopenfilenames(title="请选择需要拆分的pdf文件", filetypes=[("PDF files", "*.pdf")])
    if not pdf_paths:
        return

    for pdf_path in pdf_paths:
        corrected_name = correct_pdf_name(pdf_path)
        target_path = os.path.join(SCRIPTS_DIR, corrected_name)
        try:
            shutil.copy2(pdf_path, target_path)
            print(f"Copied {pdf_path} to {target_path}")
        except Exception as e:
            messagebox.showerror("错误", f"复制文件时出错：{str(e)}\n源文件：{pdf_path}\n目标路径：{target_path}")
            return
    
    messagebox.showinfo("成功", f"已处理 {len(pdf_paths)} 个PDF文件")

    os.startfile(SCRIPTS_DIR)

    if messagebox.askyesno("确认", "请检查PDF文件命名是否正确，与档号一致，不含其他文字、数字、符号。是否正确？"):
        if messagebox.askyesno("确认", "是否需要拆分文档？"):
            run_pdf_splitter()
    else:
        messagebox.showinfo("提示", "请手动修改PDF文件名，确认与档号一致，不含其他文字、数字、符号。")

def find_column(df, possible_names):
    for name in possible_names:
        if name in df.columns:
            return name
    return None

def parse_page_number(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        if '-' in value:
            try:
                start, _ = map(int, value.split('-'))
                return start
            except ValueError:
                print(f"Warning: Invalid page range '{value}'")
        elif value.isdigit():
            return int(value)
    print(f"Warning: Could not parse page number '{value}'")
    return None

def extract_file_info_from_sheet(df):
    file_info = {}
    
    doc_col = find_column(df, ['所属案卷档号', '档号', '案卷档号'])
    file_col = find_column(df, ['文件档号', '文件编号'])
    page_col = find_column(df, ['页号', '起始页'])
    title_col = '文件题名'
    
    print(f"Debug: Using columns: doc_col={doc_col}, file_col={file_col}, page_col={page_col}, title_col={title_col}")
    
    if not all([doc_col, file_col, page_col, title_col]):
        print("Error: Could not find all required columns")
        return None
    
    current_doc = None
    for _, row in df.iterrows():
        doc_value = str(row[doc_col])
        file_number = row[file_col]
        page_value = row[page_col]
        title = row[title_col]
        
        if doc_value.startswith('YY'):
            current_doc = doc_value
        
        if current_doc and pd.notna(file_number) and pd.notna(page_value):
            start_page = parse_page_number(page_value)
            if start_page is not None:
                if current_doc not in file_info:
                    file_info[current_doc] = []
                file_info[current_doc].append((file_number, start_page, title))
    
    for doc in file_info:
        file_info[doc] = sorted(file_info[doc], key=lambda x: x[1])
    
    return file_info

def split_pdf(input_path, output_path, start_page, end_page, total_pages):
    pdf_reader = PdfReader(input_path)
    pdf_writer = PdfWriter()
    for page_num in range(start_page - 1, min(end_page, total_pages)):
        pdf_writer.add_page(pdf_reader.pages[page_num])
    with open(output_path, 'wb') as output_file:
        pdf_writer.write(output_file)

def run_pdf_splitter():
    excel_file = os.path.join(SCRIPTS_DIR, 'input.xlsx')

    if not os.path.exists(excel_file):
        messagebox.showerror("错误", f"未找到Excel文件: {excel_file}")
        return

    try:
        xl = pd.ExcelFile(excel_file, engine='openpyxl')
        print(f"Debug: Successfully opened Excel file")
        print(f"Debug: Sheet names: {xl.sheet_names}")
    except Exception as e:
        messagebox.showerror("错误", f"打开Excel文件时出错: {str(e)}")
        return

    file_info = {}
    for sheet_name in xl.sheet_names:
        print(f"Debug: Processing sheet: {sheet_name}")
        df = xl.parse(sheet_name)
        print(f"Debug: Sheet columns: {df.columns}")
        print(df.head())
        
        sheet_info = extract_file_info_from_sheet(df)
        if sheet_info:
            file_info.update(sheet_info)
            print(f"Found file information in sheet: {sheet_name}")
        else:
            print(f"No file information found in sheet: {sheet_name}")

    print(f"Debug: Extracted file_info: {file_info}")

    if file_info:
        for filename in os.listdir(SCRIPTS_DIR):
            if filename.endswith('.pdf'):
                file_number = filename.split('.')[0]
                match = re.match(r'(YY\d+-\d+-\d+-\d+)', file_number)
                if match:
                    doc_number = match.group(1)
                else:
                    doc_number = file_number
                
                if doc_number in file_info:
                    input_path = os.path.join(SCRIPTS_DIR, filename)
                    file_ranges = file_info[doc_number]
                    print(f"Debug: Processing {doc_number}, file_ranges = {file_ranges}")

                    pdf_reader = PdfReader(input_path)
                    total_pages = len(pdf_reader.pages)

                    temp_folder = os.path.join(OUTPUT_FOLDER, doc_number)
                    os.makedirs(temp_folder, exist_ok=True)

                    for i, (sub_file_number, start_page, title) in enumerate(file_ranges):
                        if i + 1 < len(file_ranges):
                            end_page = file_ranges[i+1][1] - 1
                        else:
                            end_page = total_pages

                        output_filename = f"{sub_file_number} {title}.pdf"
                        output_path = os.path.join(temp_folder, output_filename)

                        split_pdf(input_path, output_path, start_page, end_page, total_pages)
                        print(f"Created: {output_filename} (pages {start_page}-{end_page})")

                    zip_filename = f"{doc_number}.zip"
                    zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)
                    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for root, _, files in os.walk(temp_folder):
                            for file in files:
                                zipf.write(os.path.join(root, file), file)

                    print(f"Created ZIP file: {zip_filename}")

                    shutil.rmtree(temp_folder)
                    print(f"Removed temporary PDF files for {doc_number}")

                else:
                    print(f"No matching information found for {filename}")
        
        messagebox.showinfo("成功", "PDF拆分、ZIP创建和清理完成。")
        os.startfile(OUTPUT_FOLDER)
    else:
        messagebox.showwarning("警告", "无法从Excel文件中提取文件信息。")

def open_output_folder():
    if os.path.exists(OUTPUT_FOLDER):
        os.startfile(OUTPUT_FOLDER)
    else:
        messagebox.showinfo("提示", "输出文件夹不存在。请先运行拆分文档操作。")

def process_excel_and_close(window):
    process_excel()
    window.destroy()

def process_pdf_and_close(window):
    process_pdf()
    window.destroy()

def create_document_processing_window():
    doc_window = tk.Toplevel()
    doc_window.title("文档处理")
    doc_window.geometry("300x150")

    tk.Button(doc_window, text="选择目录excel文件", 
              command=lambda: process_excel_and_close(doc_window), 
              width=20, height=2).pack(pady=10)
    tk.Button(doc_window, text="选择需要拆分的pdf文件", 
              command=lambda: process_pdf_and_close(doc_window), 
              width=20, height=2).pack(pady=10)

    # 使窗口模态，防止用户与主窗口交互
    doc_window.transient(doc_window.master)
    doc_window.grab_set()
    doc_window.master.wait_window(doc_window)

def create_gui():
    window = tk.Tk()
    window.title("PDF Splitter 工具")
    window.geometry("300x250")

    tk.Button(window, text="安装配置", command=setup_environment, width=20, height=2).pack(pady=10)
    tk.Button(window, text="文档处理", command=create_document_processing_window, width=20, height=2).pack(pady=10)
    tk.Button(window, text="拆分文档", command=run_pdf_splitter, width=20, height=2).pack(pady=10)
    tk.Button(window, text="打开文件", command=open_output_folder, width=20, height=2).pack(pady=10)

    window.mainloop()

if __name__ == "__main__":
    create_gui()