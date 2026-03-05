import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from markdownify import markdownify as md
import pdfplumber
import warnings

# 忽略一些不必要的警告（如 Excel 样式丢失警告）
warnings.filterwarnings('ignore')

# ================= 配置区域 =================

SUPPORTED_EXTENSIONS = {
    'image': ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'],
    'web':   ['.html', '.htm'],
    'csv':   ['.csv'],
    'excel': ['.xlsx', '.xls'],
    'pdf':   ['.pdf']
}

# ================= 辅助函数 =================

def get_output_path(folder_path, original_filename):
    """
    生成输出路径。
    逻辑：
    1. 默认尝试保存为 "文件名.md"
    2. 如果该文件已存在（可能是同名但不同后缀的文件生成的），则改为 "文件名_后缀.md"
    """
    base_name, ext = os.path.splitext(original_filename)
    ext_clean = ext.replace('.', '').lower()
    
    target_name = f"{base_name}.md"
    target_path = os.path.join(folder_path, target_name)
    
    # 如果文件存在，为了防止覆盖，加上后缀区分
    if os.path.exists(target_path):
        target_name = f"{base_name}_{ext_clean}.md"
        target_path = os.path.join(folder_path, target_name)
        
    return target_path

# ================= 转换逻辑函数 =================

def convert_image(file_path, folder_path):
    filename = os.path.basename(file_path)
    md_path = get_output_path(folder_path, filename)
    
    # Markdown 图片语法: ![Alt text](filename)
    content = f"# {filename}\n\n![{filename}]({filename})"
    
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(content)
    return os.path.basename(md_path)

def convert_html(file_path, folder_path):
    filename = os.path.basename(file_path)
    md_path = get_output_path(folder_path, filename)
    
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        html_content = f.read()
    
    markdown_content = md(html_content, heading_style="ATX")
    
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(markdown_content)
    return os.path.basename(md_path)

def convert_csv(file_path, folder_path):
    filename = os.path.basename(file_path)
    md_path = get_output_path(folder_path, filename)
    
    # 读取 CSV，处理空值
    df = pd.read_csv(file_path)
    df = df.fillna('') 
    md_table = df.to_markdown(index=False)
    
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(f"# {filename}\n\n")
        f.write(md_table)
    return os.path.basename(md_path)

def convert_excel(file_path, folder_path):
    filename = os.path.basename(file_path)
    md_path = get_output_path(folder_path, filename)
    
    content_list = [f"# {filename}\n"]
    
    # 读取 Excel 文件
    try:
        xls = pd.ExcelFile(file_path)
        
        # 遍历每一个 Sheet
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name)
            
            # 如果 Sheet 是空的，跳过
            if df.empty:
                continue
                
            df = df.fillna('') # 填充空值
            
            content_list.append(f"## Sheet: {sheet_name}\n")
            content_list.append(df.to_markdown(index=False))
            content_list.append("\n\n---\n") # 分隔线
            
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(content_list))
        return os.path.basename(md_path)
        
    except Exception as e:
        raise Exception(f"Excel 读取错误: {e}")

def convert_pdf(file_path, folder_path):
    filename = os.path.basename(file_path)
    md_path = get_output_path(folder_path, filename)
    
    text_content = []
    with pdfplumber.open(file_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                text_content.append(f"## 第 {i+1} 页\n\n{text}\n\n---\n")
    
    if not text_content:
        text_content.append("> ⚠️ 警告：该 PDF 未提取到文本，可能是纯图片扫描件。")

    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(f"# {filename}\n\n")
        f.write("\n".join(text_content))
    return os.path.basename(md_path)

# ================= 主程序 =================

def main():
    # 初始化 Tkinter，但不显示主窗口
    root = tk.Tk()
    root.withdraw()

    print(">>> 正在启动文件选择器...")
    
    # 弹出文件夹选择框
    folder_path = filedialog.askdirectory(title="请选择包含文件的文件夹")
    
    if not folder_path:
        print(">>> 用户取消了选择。")
        return

    print(f">>> 选中目录: {folder_path}")
    print("-" * 40)

    success_count = 0
    fail_count = 0

    # 遍历文件夹
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        
        # 跳过文件夹
        if os.path.isdir(file_path):
            continue
            
        _, ext = os.path.splitext(file_name)
        ext = ext.lower()
        
        output_name = None
        
        try:
            if ext in SUPPORTED_EXTENSIONS['image']:
                output_name = convert_image(file_path, folder_path)
                
            elif ext in SUPPORTED_EXTENSIONS['web']:
                output_name = convert_html(file_path, folder_path)
                
            elif ext in SUPPORTED_EXTENSIONS['csv']:
                output_name = convert_csv(file_path, folder_path)
                
            elif ext in SUPPORTED_EXTENSIONS['excel']:
                output_name = convert_excel(file_path, folder_path)
                
            elif ext in SUPPORTED_EXTENSIONS['pdf']:
                output_name = convert_pdf(file_path, folder_path)
            
            # 如果成功转换
            if output_name:
                print(f"[成功] {file_name} -> {output_name}")
                success_count += 1
                
        except Exception as e:
            print(f"[失败] {file_name}: {str(e)}")
            fail_count += 1

    # 结束提示
    print("-" * 40)
    result_msg = f"处理完成！\n\n成功转换: {success_count} 个\n失败文件: {fail_count} 个"
    print(result_msg)
    
    # 弹窗通知
    messagebox.showinfo("转换报告", result_msg)

if __name__ == "__main__":
    main()