import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random


class RandomQuizApp:
    def __init__(self, root):
        self.result_text = None
        self.print_button = None
        self.start_button = None
        self.num_entry = None
        self.num_label = None
        self.upload_button = None
        self.root = root
        self.root.title("随机抽题程序")

        # 文件和数据存储
        self.df = None

        # 创建GUI组件
        self.create_widgets()

    def create_widgets(self):
        # 上传文件按钮
        self.upload_button = tk.Button(self.root, text="上传Excel文件", command=self.upload_file)
        self.upload_button.pack(pady=10)

        # 输入抽题数量的标签和文本框
        self.num_label = tk.Label(self.root, text="请输入抽取题目的数量:")
        self.num_label.pack(pady=5)

        self.num_entry = tk.Entry(self.root)
        self.num_entry.pack(pady=5)

        # 开始抽题按钮
        self.start_button = tk.Button(self.root, text="开始抽题", command=self.start_draw)
        self.start_button.pack(pady=20)

        # 输出区域（显示抽取的题目）
        self.result_text = tk.Text(self.root, height=10, width=50)
        self.result_text.pack(pady=10)

        # 提供打印按钮
        self.print_button = tk.Button(self.root, text="打印题目", command=self.print_quiz)
        self.print_button.pack(pady=10)

    def upload_file(self):
        """上传Excel文件并读取数据"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            try:
                # 读取Excel文件
                self.df = pd.read_excel(file_path)
                # 确保文件中有题目列
                if '知识点' not in self.df.columns:
                    messagebox.showerror("错误", "Excel文件中没有'知识点'列！")
                    return
                messagebox.showinfo("成功", "文件上传成功！")
            except Exception as e:
                messagebox.showerror("错误", f"文件读取失败: {e}")

    def start_draw(self):
        """开始抽题"""
        if self.df is None:
            messagebox.showerror("错误", "请先上传一个Excel文件！")
            return

        # 获取用户输入的抽取数量
        try:
            num_questions = int(self.num_entry.get())
            if num_questions <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("错误", "请输入一个有效的抽题数量！")
            return

        # 随机抽取题目
        questions = self.df['知识点'].dropna().tolist()  # 确保去掉空值
        if num_questions > len(questions):
            messagebox.showerror("错误", "抽题数量不能大于题目总数！")
            return

        selected_questions = random.sample(questions, num_questions)

        # 显示结果
        self.result_text.delete(1.0, tk.END)  # 清空现有的内容
        for idx, question in enumerate(selected_questions, start=1):
            self.result_text.insert(tk.END, f"题目 {idx}: {question}\n")

    def print_quiz(self):
        """打印抽取的题目"""
        quiz_content = self.result_text.get(1.0, tk.END).strip()
        if not quiz_content:
            messagebox.showerror("错误", "没有题目可以打印！")
            return

        # 打印功能（简单地将内容复制到剪贴板或打印）
        try:
            with open("抽题结果.txt", "w") as f:
                f.write(quiz_content)
            messagebox.showinfo("成功", "题目已保存为'抽题结果.txt'，可以打印。")
        except Exception as e:
            messagebox.showerror("错误", f"保存文件失败: {e}")


# 创建主窗口
root = tk.Tk()
app = RandomQuizApp(root)
root.mainloop()

if __name__ == '__main__':
    root.mainloop()