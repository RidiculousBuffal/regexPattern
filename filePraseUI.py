import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import font as tkFont


def process_str_to_list(input_string):
    output_string = re.sub(r'\s+', ' ', input_string)
    items = output_string.split("#题目")
    items = [("#题目 " + item.strip()) for item in items if item.strip()]
    dic = {}

    for item in items:
        pattern = r'#题目\s+(.*?)\s+问题.*?用户填写\s+(.*?)(?=\s*#|$)'
        matches = re.findall(pattern, item)

        for match in matches:
            dic[match[0]] = match[1]
    return dic


def process_file(file_path):
    df = pd.read_excel(file_path)
    ls = []

    for i in range(0, len(df)):
        dic_ = {}
        result = ' '.join(df.iloc[i, 5:].astype(str)).replace('NaN', '').replace("nan", '')
        dic_.update(dict(df.iloc[i, 0:4]))
        dic_.update(process_str_to_list(result))
        ls.append(dic_)

    df_new = pd.DataFrame(ls)
    return df_new


def upload_and_process():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        try:
            processed_df = process_file(file_path)
            messagebox.showinfo("Success", "File uploaded and processed successfully!")
            save_file(processed_df)
        except Exception as e:
            messagebox.showerror("Error", str(e))


def save_file(df):
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if save_path:
        try:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", "File saved successfully at:\n" + save_path)
        except Exception as e:
            messagebox.showerror("Error", str(e))


# 创建主窗口
root = tk.Tk()
root.title("Excel Processor")

# 设置窗口大小
root.geometry("600x400")

# 计算中心位置
window_width = 600
window_height = 400
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# 设置字体
font_style = tkFont.Font(family="Times New Roman", size=12)

# 创建操作简介标签
intro_label = tk.Label(root,
                       text="操作简介：\n1. 点击 'Upload Excel File' 上传Excel文件。\n2. 处理完成后选择保存文件的位置和名称。",
                       font=font_style)
intro_label.pack(pady=10)

# 创建上传按钮
upload_button = tk.Button(root, text="Upload Excel File", command=upload_and_process, font=font_style)
upload_button.pack(pady=20)

# 运行主循环
root.mainloop()
