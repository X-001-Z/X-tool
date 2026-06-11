# PDF2PPT

PDF2PPT 是一个 Windows 桌面工具，可将 PDF 的每一页转换成图片并写入 PowerPoint 文件。转换完全在本机进行，不需要安装 Microsoft Office，也不会上传 PDF。

## 主要功能

- 批量转换多个 PDF
- 屏幕、平衡、高清三档质量预设
- 支持自定义 DPI 和 JPEG 质量
- 支持密码保护的 PDF
- 支持取消转换
- 混合横竖页面等比居中，避免拉伸
- 同名 PPTX 自动添加序号，不覆盖现有文件
- 超大页面自动降低实际 DPI，避免内存占用失控
- 图形界面和命令行模式

## 下载与使用

[前往 GitHub Releases 下载最新版本](https://github.com/X-001-Z/X-tool/releases/latest)

下载 Windows 64 位便携版，解压后运行 `PDF2PPT.exe`。

1. 添加一个或多个 PDF。
2. 选择质量预设。
3. 可选设置统一输出目录。
4. 点击“开始转换”。

程序生成的是图片型 PPTX。页面视觉效果会保留，但文字、表格和图形不能在 PowerPoint 中单独编辑。

### 质量预设

| 预设 | DPI | JPEG 质量 | 适用场景 |
|---|---:|---:|---|
| 屏幕 | 150 | 85 | 预览、分享、较小文件 |
| 平衡 | 200 | 88 | 日常使用，默认推荐 |
| 高清 | 300 | 92 | 打印或大屏展示 |

### 命令行

```powershell
PDF2PPT.exe --convert input.pdf --dpi 200 --quality 88
PDF2PPT.exe --convert a.pdf b.pdf --output-dir output
PDF2PPT.exe --convert protected.pdf --password "your-password"
```

查看全部参数：

```powershell
PDF2PPT.exe --help
```

## 系统要求

- Windows 10 或 Windows 11，64 位
- 不需要安装 Python、Microsoft Office 或 Adobe Acrobat

## 隐私

PDF2PPT 不包含联网、遥测、广告或自动上传功能。详细说明见 [PRIVACY.md](PRIVACY.md)。

## 从源码构建

要求：Windows 10/11、Python 3.12、PowerShell。

```powershell
git clone https://github.com/X-001-Z/X-tool.git
cd X-tool
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

脚本会创建独立的 `.venv-build` 环境、安装锁定版本的依赖、执行测试，并在 `outputs` 中生成 EXE。

生成便携版发布包：

```powershell
powershell -ExecutionPolicy Bypass -File .\release.ps1
```

## 测试

```powershell
.\.venv-build\Scripts\python.exe -m unittest -v test_pdf2ppt.py
```

测试覆盖基本转换、DPI、混合页面尺寸、密码 PDF、像素限制、取消及同名文件保护。

## 已知限制

- 输出内容为图片，不可逐元素编辑。
- PowerPoint 文件只能使用一种幻灯片尺寸；混合尺寸 PDF 以第一页为基准，其他页面等比居中。
- 当前公开构建未进行商业代码签名，Windows 首次运行时可能显示“未知发布者”提示。
- PDF 中原有的色彩配置和深色背景会按 PDFium 的渲染结果输出。

## 开源许可

项目源码使用 [MIT License](LICENSE)。软件包含其他开源组件，其许可信息见 [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) 和发布包中的 `licenses` 目录。

## 安全问题

请不要在公开 issue 中提交包含敏感 PDF、密码或个人信息的样本。报告方式见 [SECURITY.md](SECURITY.md)。
