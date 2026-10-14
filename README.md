# PDF 工具箱 · 白板版

Windows 本地 PDF 文档处理工具：合并、拆分、删页、抽页、转换、水印一站式完成。所有文件均在本地处理，不需要账号登录，也不上传用户文档。

## 功能特性

- **PDF 合并**：按文件列表顺序把多个 PDF 合成为一个
- **PDF 拆分**：按 `1、3-5` 等页码范围生成一个或多个拆分文件
- **删除页面 / 提取页面**：按页码删除，或把指定页面复制到新文件
- **转换为图片**：按设定分辨率把 PDF 页面输出为 PNG
- **转换为 Word**：提取页面文字和图片生成 DOCX 文档
- **PDF 与 Word 合并**：Word 先通过本机 Microsoft Word 转换为 PDF，再按顺序统一合并
- **文字水印**：可调文字、颜色、透明度、旋转角度、位置和缩放比例
- **图片水印**：支持 PNG、JPG、WebP 等图片水印并调节比例
- **拖放排序**：文件列表支持批量添加、拖放调整顺序和移除

## 下载

单文件安装程序见 [Releases 页面](https://github.com/shuhan-AFK/pdf-toolbox/releases)，无需安装、双击即用。

- 适用平台：Windows 10 / 11 64 位
- 涉及 Word 的功能需要本机已安装 Microsoft Word

## 工程结构

```text
.
├── app.py                     界面适配与启动器（可读源码）
├── app.marshalled             核心处理逻辑（编译资源，随包交付）
├── make_icon.py               应用图标生成脚本
├── PDF工具箱_白板版.spec       PyInstaller 打包配置
├── build.ps1                  一键构建脚本
├── assets/                    图标资源
└── whiteboard_preview.png     界面预览
```

说明：界面使用通用白色主题与项目自有图标，不包含任何第三方素材。核心处理逻辑以编译资源 `app.marshalled` 随包交付（工程未保留其原始可读源码）。

## 技术栈

Python 3.11 · PyQt5 · PyMuPDF · Pillow · python-docx · pywin32（Word COM 自动化）· PyInstaller

## 构建

环境要求：Python 3.11（64 位），依赖 `PyQt5 PyMuPDF Pillow python-docx pywin32 pyinstaller`

```powershell
.\build.ps1
```

脚本会先生成图标，再调用 PyInstaller 打包，输出到 `dist/PDF工具箱_白板版.exe`。

## 界面预览

![界面预览](whiteboard_preview.png)

## 版权

© 2026 shuhan-AFK。保留所有权利。
