import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import random

class DataSamplerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("数据")
        self.root.geometry("500x280")

        self.df_original = None
        self.df_result = None

        self.btn_load = tk.Button(root, text="导入 Excel 文件", command=self.load_excel, width=20, height=2)
        self.btn_load.pack(pady=10)

        self.lbl_file = tk.Label(root, text="尚未选择文件", fg="gray")
        self.lbl_file.pack()

        self.btn_process = tk.Button(root, text="排查数据", command=self.process_data, width=20, height=2, state="disabled")
        self.btn_process.pack(pady=10)

        self.lbl_info = tk.Label(root, text="")
        self.lbl_info.pack()

        self.btn_export = tk.Button(root, text="导出结果", command=self.export_result, width=20, height=2, state="disabled")
        self.btn_export.pack(pady=10)

    def load_excel(self):
        file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if not file_path:
            return

        try:
            self.df_original = pd.read_excel(file_path, header=None)
            if self.df_original.shape[1] < 4:
                messagebox.showerror("错误", "数据至少需要 4 列（时间、代码、数值1、数值2、数值3）")
                return

            time_col = self.df_original[0]

            if pd.api.types.is_datetime64_any_dtype(time_col):
                parsed = time_col
            else:
                time_str = time_col.astype(str).str.strip()
                formats = [
                    '%Y/%m/%d %H:%M:%S',
                    '%Y-%m-%d %H:%M:%S',
                    '%Y/%m/%d %H:%M',
                    '%H:%M:%S',
                    '%H:%M',
                ]
                parsed = None
                for fmt in formats:
                    candidate = pd.to_datetime(time_str, format=fmt, errors='coerce')
                    if not candidate.isna().all():
                        parsed = candidate
                        break

                if parsed is None:
                    sample = time_str.head(3).tolist()
                    messagebox.showerror(
                        "错误",
                        f"第一列无法解析为时间。\n前几行数据示例：{sample}\n"
                        "请确保格式如：2026/6/9 16:42:53 或 16:42:53"
                    )
                    return

            self.df_original[0] = parsed
            self.lbl_file.config(text=f"已加载：{os.path.basename(file_path)}，共 {len(self.df_original)} 行")
            self.btn_process.config(state="normal")
            self.df_result = None
            self.lbl_info.config(text="")
            self.btn_export.config(state="disabled")
        except Exception as e:
            messagebox.showerror("读取失败", f"读取 Excel 文件出错：{str(e)}")

    def process_data(self):
        if self.df_original is None:
            messagebox.showwarning("提示", "请先导入 Excel 文件")
            return

        df = self.df_original.copy()
        df['__index__'] = range(len(df))

        df['__second__'] = df[0].dt.floor('s')
        groups = df.groupby('__second__', sort=False)

        result_parts = []
        for _, group in groups:
            group_size = len(group)
            if group_size <= 20:
                continue
            kept = group.iloc[::2]
            if len(kept) > 20:
                kept = kept.sample(n=20, random_state=random.randint(0, 10000)).sort_index()
            result_parts.append(kept)

        if not result_parts:
            messagebox.showinfo("处理结果", "处理后无保留数据（所有秒的数据均不足或等于20行，已被删除）")
            self.df_result = None
            self.lbl_info.config(text="无数据保留")
            self.btn_export.config(state="disabled")
            return

        self.df_result = pd.concat(result_parts).sort_values('__index__')
        self.df_result = self.df_result.drop(columns=['__second__', '__index__'])
        self.df_result[0] = self.df_result[0].dt.strftime('%H:%M:%S')

        self.lbl_info.config(text=f"处理完成，保留 {len(self.df_result)} 行数据")
        self.btn_export.config(state="normal")
        messagebox.showinfo("处理完成", f"数据抽选完成！\n保留 {len(self.df_result)} 行数据，可点击“导出结果”保存。")

    def export_result(self):
        if self.df_result is None:
            messagebox.showwarning("提示", "没有可导出的数据")
            return

        file_path = filedialog.asksaveasfilename(
            title="保存结果",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if not file_path:
            return

        try:
            self.df_result.to_excel(file_path, index=False, header=False)
            messagebox.showinfo("成功", f"结果已保存至：\n{file_path}")
        except Exception as e:
            messagebox.showerror("保存失败", f"保存文件时出错：{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DataSamplerApp(root)
    root.mainloop()
