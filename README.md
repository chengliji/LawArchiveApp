# 律师案卷归档小工具

一个为律师设计的自定义案卷归档管理工具，能将办案过程中形成的文件（如 .doc .pdf .jpg .png .txt .ppt .tiff等）一键生成含有归档目录的 pdf 文件，方便律师进行电子归档。

[![Python](https://img.shields.io/badge/Python-100%25-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

## 一、功能特性

- 可以自定义案卷内容和案卷顺序。可以根据不同地区不同律所的特点，自定义案卷内容和案卷顺序。
- 电子书签（侧边栏导航）。生成的电子文件 PDF 能自带左侧大纲导航。
- 卷宗自动添加页码。卷宗每页自动添加页码，页码会索引至自动生成的目录页中。
- 电子卷宗瘦身（网上立案压缩模式）。开启电子卷宗瘦身的勾选项，程序会在后台启动画质压缩算法。
- 分屏布局：左侧是你的案件目录，右侧是多功能仪表盘。
- 实时预览窗：点击左侧列表里的图片或 PDF，右侧会立刻渲染出缩略图（Word/Excel 暂不支持实时渲染，会显示提示图标）。
- 质检与日志面板：合并时不再是干等，右侧会有滚动的文字告诉你进行到哪一步了。同时会自动统计总文件数。
- 进度草稿箱：在界面最顶部加上了“文件”菜单栏，下班前没弄完，直接“保存草稿(.json)”，第二天“加载草稿”瞬间恢复所有文件和顺序！
- 自定义防伪水印：点击合并前，会弹窗问你是否需要打水印。输入“XX案件归档”，生成的PDF每一页都会被打上倾斜的灰色防伪底纹。
- 支持的文件格式：*.pdf *.doc *.docx *.xls *.xlsx *.ppt *.pptx *.txt *.rtf *.jpg *.jpeg *.png *.bmp *.tif *.tiff
- 支持原生文件拖拽上传。
- 支持双击列表中的文件调用系统软件直接打开查阅。

## 二、软件界面

  <img width="1502" height="1039" alt="屏幕截图 2026-05-15 154112" src="https://github.com/user-attachments/assets/450bca14-a5da-4150-a1bd-a982b4eeeddc" />


## 三、快速开始

### 1.环境要求

- Python 3.7+

### 2.安装（Windows 平台）

##### 第一步：在 Windows 上准备环境

在准备打包的 Windows 电脑上，打开命令提示符（CMD），安装依赖库：

```pip install PySide6 PyMuPDF reportlab pillow```

#### 第二步：执行打包命令

在 CMD 中，使用 cd 命令切换到你存放 ArchiveApp.py 的目录。然后，输入以下这行命令并回车：

```python -m PyInstaller --noconsole --onefile ArchiveApp.py```

(💡 进阶小贴士：如果你想给软件加个好看的图标，可以准备一个 .ico 格式的图片放同目录下，把命令改成 pyinstaller --noconsole --onefile --icon=logo.ico ArchiveApp.py)

#### 第三步：获取你的 .exe

在项目代码目录中，打开 dist 文件夹，里面的 ArchiveApp.exe 就是你最终的劳动成果！
