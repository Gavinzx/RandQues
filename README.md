# 随机抽题程序
该程序是基于 Windows 操作系统开发的桌面应用程序，用于从 Excel 文件中随机抽取题目并生成题目列表，用户可以上传一个包含知识点的 Excel 文件，指定抽取题目的数量，并生成可供打印的题目列表。使用 Python 和 Tkinter 库构建 GUI 界面，并通过 pandas 和 openpyxl 处理 Excel 文件。程序打包后可生成一个 .exe 文件，可以直接在 Windows 环境下运行，无需安装 Python。

## 功能
+ 上传 Excel 文件：支持上传包含知识点的 Excel 文件，程序会自动读取其中的题目。
+ 随机抽取题目：用户可以指定要抽取的题目数量，程序会随机从 Excel 文件中抽取相应数量的题目。
+ 保存抽题结果：用户可以将抽取的题目保存到本地，生成一个可打印的文本文件。
+ 界面友好：采用 Tkinter 库，提供简洁直观的用户界面，适合初学者和日常使用。

## 特性

+ 支持 .xlsx 格式的 Excel 文件。
+ 抽题数量支持用户自定义。
+ 题目保存为文本文件，方便打印。
+ 错误信息和操作反馈直接显示在应用内，无需依赖外部弹窗。

## 安装与使用

1. 克隆仓库
    ```powershell
    git clone https://github.com/Gavinzx/RandQues.git
    cd RandQues
    ```
2. 安装依赖 

    首先，确保你已经安装了 Python 3.10 和 pip。然后通过以下命令安装依赖：

    ```powershell
    pip install -r requirements.txt
    ```
3. 运行程序

    ```powershell
    python main.py
    ```

4. 使用

    1. 打开程序后，点击“上传Excel文件”按钮，选择包含题目的 Excel 文件，仓库中 `static` 目录下有 Excel 模板文件。

    2. 输入要抽取的题目数量，然后点击“开始抽题”按钮，程序会随机抽取指定数量的题目并展示。

    3. 如果需要保存题目，点击“保存题目”按钮选择保存路径。

## 使用 PyInstaller 打包成 EXE 文件

如果你需要将程序打包成可执行的 .exe 文件，可以使用 PyInstaller 工具。具体步骤如下：

1. 安装 PyInstaller

   如果你还没有安装 PyInstaller，可以通过以下命令安装：

   ```powershell
   pip install pyinstaller
   ```

2. 打包程序

   在项目根目录下，你可以执行 build.ps1 脚本来打包程序。确保你已经安装了 PowerShell 并配置好环境。执行脚本时，它会使用 PyInstaller 打包程序并生成 .exe 文件。

   运行以下命令来执行脚本：

   ```powershell
   .\build.ps1
   ```

3. 打包完成后 .exe 文件会生成在 dist 文件夹下，你可以直接运行这个 .exe 文件，而无需安装 Python 环境。

## 技术栈
Python 3.10+

Tkinter（GUI 框架）

pandas（处理 Excel 文件）

## 许可证

本项目采用 MIT 许可证，详情请见 LICENSE。

