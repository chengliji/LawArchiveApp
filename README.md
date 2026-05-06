# 律所案卷归档小工具

一个为律所（律师）设计的自定义案卷归档管理工具，能将办案过程中形成的文件（如 .doc .pdf .jpg .png .txt .ppt .tiff等）一键生成含有归档目录的 pdf 文件，方便律所进行电子归档。

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

## 二、软件界面

  <img width="1888" height="1002" alt="屏幕截图 2026-05-04 114447" src="https://github.com/user-attachments/assets/a0445e9d-b365-4a75-ae9b-3a0bb0e1974b" />


## 三、快速开始

### 1.环境要求

- Python 3.7+

### 2.安装（Windows 平台）

##### 第一步：在 Windows 上准备环境

在准备打包的 Windows 电脑上，打开命令提示符（CMD），安装依赖库：

```pip install pywin32 PyMuPDF reportlab pillow pyinstaller```

#### 第二步：执行打包命令

在 CMD 中，使用 cd 命令切换到你存放 ArchiveApp.py 的目录。然后，输入以下这行命令并回车：

```python -m PyInstaller --noconsole --onefile ArchiveApp.py```

(💡 进阶小贴士：如果你想给软件加个好看的图标，可以准备一个 .ico 格式的图片放同目录下，把命令改成 pyinstaller --noconsole --onefile --icon=logo.ico ArchiveApp.py)

#### 第三步：获取你的 .exe

在项目代码目录中，打开 dist 文件夹，里面的 ArchiveApp.exe 就是你最终的劳动成果！

### 3.安装（macOS 平台）

#### 第一步：适配苹果字体与系统

打开 ArchiveApp.py，找到最前面的这段代码：

```Python
def get_windows_font():
    if not IS_WINDOWS: return None
    font_paths = [
        "C:\\Windows\\Fonts\\simsun.ttc",
        "C:\\Windows\\Fonts\\msyh.ttc",
        "C:\\Windows\\Fonts\\simhei.ttf"
    ]
    for path in font_paths:
        if os.path.exists(path): return path
    return None

SYS_FONT_PATH = get_windows_font()
SYS_FONT_PATH = get_windows_font()
```

将它替换为跨平台版本：

```Python
def get_system_font():
    font_paths = []
    if sys.platform.startswith('win'):
        font_paths = [
            "C:\\Windows\\Fonts\\simsun.ttc",
            "C:\\Windows\\Fonts\\msyh.ttc",
            "C:\\Windows\\Fonts\\simhei.ttf"
        ]
    elif sys.platform == 'darwin': # macOS 系统
        font_paths = [
            "/System/Library/Fonts/PingFang.ttc",       # 苹方字体
            "/System/Library/Fonts/STHeiti Light.ttc",  # 华文黑体
            "/Library/Fonts/Arial Unicode.ttf"
        ]
    
    for path in font_paths:
        if os.path.exists(path): return path
    return None

SYS_FONT_PATH = get_system_font()
```

#### 第二步：在 Mac 电脑上准备环境

打开 Mac 的 终端 (Terminal)，安装和之前一样的依赖库：

```pip3 install PyMuPDF reportlab pillow customtkinter pyinstaller```

#### 第三步：执行 Mac 专属打包命令

在终端中使用 cd 命令进入你存放代码的文件夹。然后执行：

```pyinstaller --windowed --onefile ArchiveApp.py```

#### 第四步：收获你的 Mac 软件

打包完成后，进入 dist 文件夹，你会看到一个名为 ArchiveApp.app 的文件，它带着标准 Mac 软件的图标。你可以把它拖到“应用程序”文件夹里使用。

注意：当你双击运行它时，Mac 大概率会弹窗提示：“ArchiveApp.app 已损坏，无法打开。你应该将它移到废纸篓。”时，有两种解决方法：

1 打开 Mac 终端，输入以下命令（注意最后有一个空格），然后把你的 .app 文件拖进终端里，按下回车即可解除锁定：

```sudo xattr -rd com.apple.quarantine ```

(例如：sudo xattr -rd com.apple.quarantine /Users/name/Desktop/dist/ArchiveApp.app)

2 右键强开法：

按住键盘上的 Control 键，鼠标点击该软件，在弹出的菜单中选择“打开”。系统会再次警告你，但这次会多出一个“打开”按钮供你强行放行。
