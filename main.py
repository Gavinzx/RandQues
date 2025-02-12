import tkinter as tk
from tkinter import filedialog
import pandas as pd
import random

AUTHOR = "Gavin"
VERSION = "@VERSION@"
DATE = "@DATE@"


class RandomQuizApp:
    def __init__(self, root):
        self.signature_label = None
        self.message_label = None
        self.save_button = None
        self.result_text = None
        self.start_button = None
        self.num_entry = None
        self.num_label = None
        self.upload_button = None
        self.root = root
        self.root.title("随机抽题程序")

        # 设置窗口大小和最小大小
        self.root.geometry("700x500")  # 初始窗口大小
        self.root.minsize(700, 500)  # 最小窗口大小

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
        self.result_text = tk.Text(self.root, height=12, width=60)  # 调整文本框大小
        self.result_text.pack(padx=10, pady=10, fill='both', expand=True)  # 使文本框随窗口大小调整

        # 提供保存按钮
        self.save_button = tk.Button(self.root, text="保存题目", command=self.save_quiz)
        self.save_button.pack(pady=10)

        # 消息显示区域（用于显示错误和成功信息）
        self.message_label = tk.Label(self.root, text="", fg="red", wraplength=650)
        self.message_label.pack(pady=5)

        # 签名行（显示程序开发者和版本信息）
        self.signature_label = tk.Label(self.root, text=f"{AUTHOR} | {VERSION} | {DATE}", fg="gray", font=("Arial", 8))
        self.signature_label.pack(side="bottom", pady=10)

    def upload_file(self):
        """上传Excel文件并读取数据"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            try:
                # 读取Excel文件
                self.df = pd.read_excel(file_path)
                # 确保文件中有题目列
                if '知识点' not in self.df.columns:
                    self.show_message("错误: Excel文件中没有'知识点'列！", "red")
                    return
                self.show_message("文件上传成功！", "green")
            except Exception as e:
                self.show_message(f"错误: {e}", "red")

    def start_draw(self):
        """开始抽题"""
        if self.df is None:
            self.show_message("错误: 请先上传一个Excel文件！", "red")
            return

        # 获取用户输入的抽取数量
        try:
            num_questions = int(self.num_entry.get())
            if num_questions <= 0:
                raise ValueError
        except ValueError:
            self.show_message("错误: 请输入一个有效的抽题数量！", "red")
            return

        # 随机抽取题目
        questions = self.df['知识点'].dropna().tolist()  # 确保去掉空值
        if num_questions > len(questions):
            self.show_message("错误: 抽题数量不能大于题目总数！", "red")
            return

        selected_questions = random.sample(questions, num_questions)

        # 显示结果
        self.result_text.delete(1.0, tk.END)  # 清空现有的内容
        for idx, question in enumerate(selected_questions, start=1):
            self.result_text.insert(tk.END, f"{idx}: {question}\n")

        self.show_message("抽题成功！", "green")

    def save_quiz(self):
        """保存抽取的题目到指定路径"""
        quiz_content = self.result_text.get(1.0, tk.END).strip()
        if not quiz_content:
            self.show_message("错误: 没有题目可以保存！", "red")
            return

        # 打开文件保存对话框，选择保存路径和文件名
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(quiz_content)
                self.show_message(f"题目已保存到: {file_path}", "green")
            except Exception as e:
                self.show_message(f"错误: 保存文件失败: {e}", "red")

    def show_message(self, message, color):
        """更新消息显示区域"""
        self.message_label.config(text=message, fg=color)


# 创建主窗口
root = tk.Tk()
app = RandomQuizApp(root)
root.mainloop()

if __name__ == '__main__':
    root.mainloop()
