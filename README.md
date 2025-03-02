# Word2PDF 转换器

一个简单易用的Word文档批量转换PDF工具，基于PyQt6开发的图形界面应用。

## 功能特点

- 批量转换：支持同时转换多个Word文档
- 暂停/继续：可随时暂停或继续转换过程
- 进度显示：实时显示转换进度和当前处理文件
- 优雅的界面：现代化的UI设计，操作简单直观

## 系统要求

- macOS操作系统
- Microsoft Word
- Python 3.9+

## 安装使用

1. 克隆仓库到本地
```bash
git clone [仓库地址]
cd word2pdf
```

2. 创建虚拟环境并安装依赖
```bash
python -m venv venv
source venv/bin/activate  # Windows使用: .\venv\Scripts\activate
pip install -r requirements.txt
```

3. 运行应用
```bash
python word2pdf_app.py
```

## 使用方法

1. 点击"选择输入文件夹"按钮选择包含Word文档的文件夹
2. 点击"选择输出文件夹"按钮选择PDF文件的保存位置
3. 点击"开始转换"按钮开始转换过程
4. 可以使用"暂停"和"停止"按钮控制转换过程

## 许可证

MIT License