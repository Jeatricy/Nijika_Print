# Nijika Print

**1.0.1 · Windows 本地批量打印桌面应用**

基于 WebView2 的独立桌面窗口。文件在本机处理，不需要打开网站。

## 下载与运行

从 GitHub Releases 下载 `Nijika_Print-1.0.1-windows-x64.exe`，双击运行。

- Windows 10/11 x64，需安装 Microsoft Edge WebView2 Runtime。
- Release 为单文件程序，包含 Python 运行组件和 SumatraPDF 3.6.1；首次启动会解压运行资源到临时目录。
- Office 文档打印仍需本机安装相应 Microsoft Office / WPS。
- 程序未做商业代码签名；请核对 Release 中提供的 SHA-256 文件。

## 功能

- Word、Excel、PowerPoint、PDF 及常见图片批量打印。
- 文件夹浏览、自然排序、文件名/扩展名筛选、拖放导入、去重。
- 独立勾选与高亮多选，支持调整打印顺序。
- 打印机、彩色/黑白、单双面、PDF 拼版（每张 1～16 页）及 Excel 转 PDF 模式。
- 左栏文件工具栏、短控件与行内标签、紧凑文件行；右侧队列只显示文件名等信息，不显示路径副行。
- 打印前确认、进度状态、取消后续提交，以及退出前清理。

移除或清空队列只影响列表，不会删除原文件。开始打印时会对队列和设置生成快照，之后的界面修改不会改变当前任务。

“取消打印”只停止后续提交；已经进入 Windows 打印队列的任务需在系统中取消。

## 源码运行

建议使用 Python 3.11：

```powershell
python -m pip install -r requirements.txt
python -m nijika_print
```

源码仓库不包含 EXE、旧版本、批处理启动脚本或测试样例；根目录的 `icon.ico` 是正式应用图标，用于 EXE、任务栏和标题栏。打印引擎已集成在 `nijika_print/engine.py`，不依赖任何旧版源码。

源码运行时需要单独安装 SumatraPDF。程序会从常见安装目录和 PATH 查找，也可以指定：

```powershell
$env:NIJIKA_SUMATRA_PATH = 'C:\Program Files\SumatraPDF\SumatraPDF.exe'
python -m nijika_print
```

缺少 SumatraPDF 时，Excel 转 PDF、图片和静默 PDF 打印不能按完整流程工作。

## 项目结构

```text
nijika_print/
  __main__.py       Python 模块入口
  app.py            桌面窗口、原生拖放与诊断
  backend.py        文件、打印队列和任务状态
  engine.py         独立打印引擎
  branding.py       名称与版本
  resources.py      应用图标资源定位
  web/              随程序提供的界面源码
packaging/
  version_info.txt  Windows 文件版本元数据
Nijika_Print.spec   单文件 Release 构建配置
```

## 构建 Release

安装依赖和 PyInstaller，并准备外部的 SumatraPDF 3.6.1 可执行文件。它仅作为打包输入，不提交到 Git：

```powershell
python -m pip install -r requirements.txt "pyinstaller>=6,<7"
$env:NIJIKA_SUMATRA_BUILD_PATH = 'C:\Program Files\SumatraPDF\SumatraPDF.exe'
python -m PyInstaller --noconfirm Nijika_Print.spec
```

产物：`dist/Nijika_Print-1.0.1-windows-x64.exe`。

可在不创建窗口、不提交打印的情况下运行依赖和 PDF 拼版诊断：

```powershell
python -m nijika_print --diagnose "$env:TEMP\nijika-diagnostics.json"
```

该参数也适用于 Release EXE。启动失败日志写入 `%LOCALAPPDATA%\Nijika Print\logs\startup-error.log`。

## 版本与第三方组件

1.0.1 是图标修正版，统一使用根目录 `icon.ico`，其余界面和打印功能保持不变。详见 `CHANGELOG.md` 和 `THIRD_PARTY_NOTICES.md`。

本仓库为公开仓库，未额外指定项目代码的开源许可。第三方组件保留其各自许可及权利声明。

