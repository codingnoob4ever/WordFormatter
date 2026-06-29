# Word 格式修改器

一键将 Markdown / TXT / DOCX 文件排版为党政公文标准格式。

**离线运行，不联网，不调用 AI，完全本地处理。**

[![macOS](https://img.shields.io/badge/macOS-✓-blue)](https://github.com/codingnoob4ever/WordFormatter)
[![Windows](https://img.shields.io/badge/Windows-✓-blue)](https://github.com/codingnoob4ever/WordFormatter)
[![Python](https://img.shields.io/badge/Python-3.9+-green)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-yellow)](https://github.com/codingnoob4ever/WordFormatter/blob/main/LICENSE)

---

## 这是什么？

一个本地网页工具，帮你在浏览器里把 Word 文档自动排版成规范的公文格式。

**典型场景：**
- 把 AI 生成的 Markdown 讲稿转成标准 Word 文档
- 把录音转文字稿统一排版后交付
- 批量处理多个文档，一键套用相同格式

**不做什么：**
- 不联网上传文件（所有处理在本地完成）
- 不修改原文（输出新文件，原文件不动）
- 不调用任何 AI 接口

---

## 快速开始

### 第一步：安装依赖（只需一次）

| 平台 | 操作 |
|------|------|
| **macOS** | 双击 `install.command` |
| **Windows** | 双击 `install.bat` |

> 前提：电脑上已安装 Python 3.9+。如果没有，先去 [python.org](https://www.python.org/downloads/) 下载安装。

### 第二步：启动程序

| 平台 | 操作 |
|------|------|
| **macOS** | 双击 `start.command` |
| **Windows** | 双击 `start.bat` |

浏览器会自动打开 `http://localhost:8080`，这就是工具界面。

### 第三步：开始使用

1. 选择预设「党政公文标准」
2. 拖入你的 `.docx` / `.txt` / `.md` 文件
3. 点「扫描文档」，确认标题识别结果
4. 点「确认格式化」，等待处理完成
5. 下载格式化的文件

---

## 内置预设

### 党政公文标准

| 要素 | 字体 | 字号 | 格式 |
|------|------|------|------|
| 文章标题 | 方正小标宋简体 | 二号（22pt） | 加粗、居中 |
| 一级标题 | 黑体 | 三号（17pt） | 左对齐 |
| 二级标题 | 楷体 | 三号（17pt） | 左对齐 |
| 三级标题 | 仿宋 | 三号（17pt） | 加粗、左对齐 |
| 正文 | 仿宋 | 三号（17pt） | 首行缩进 2 字符，1.34 倍行距 |
| 作者署名 | 楷体 | 四号（14pt） | 居中 |

页边距：上 3.8 / 下 3.2 / 左 2.7 / 右 2.7 cm

标题编号：一、（一）、1.

### 通用学术格式

适用于论文、报告等场景。

---

## 使用技巧

### 字体设置

如果你的电脑没有方正小标宋简体等字体，可以在工具右上角的「字体设置」里修改字体映射。工具会自动扫描你系统里已安装的中文字体。

### 自定义预设

在界面里调整参数后，点「保存当前预设」或「另存为新预设」，就能创建自己的格式规范。预设保存在 `config.json` 里，也可以直接编辑。

### 跳过预览

处理大量文件时，可以勾选「跳过预览确认」，上传后自动处理，无需逐份确认。

### 页码更新

格式化后的文档页码是 Word 域。打开文件后按 `Ctrl+A` 再按 `F9` 即可刷新页码。

---

## 常见问题

### Q: 启动时报错「No module named 'docx'」？

依赖没装。双击 `install.command`（Mac）或 `install.bat`（Windows）安装。

### Q: 字体下拉框里没有我想要的字体？

工具只显示系统中已安装的中文字体。如果需要特定字体，先把字体文件安装到系统里。

### Q: 格式化后字体不对，显示为宋体？

Word 回退到了默认字体。说明你电脑上没有预设中指定的字体。在界面「字体设置」里，把字体名改成你电脑上有的。

### Q: 支持 .doc 格式吗？

不支持。请在 Word 里另存为 `.docx` 后再使用。

### Q: 文件安全吗？会传到网上吗？

**不会。** 所有处理在你电脑本地完成，不需要联网。上传的文件存在项目目录的 `uploads/` 里，处理完的放在 `results/` 里，关闭程序后可随时删除。

---

## 项目结构

```
WordFormatter/
├── server.py          ← Flask 后端服务
├── formatter.py       ← 格式处理引擎
├── index.html         ← 网页前端界面
├── config.json        ← 格式预设配置
├── requirements.txt   ← Python 依赖清单
├── install.command    ← Mac 依赖安装（双击）
├── install.bat        ← Windows 依赖安装（双击）
├── start.command      ← Mac 启动脚本（双击）
├── start.bat          ← Windows 启动脚本（双击）
└── README.md
```

---

## 技术栈

- Python 3.9+
- Flask（Web 服务）
- python-docx（Word 文档读写）
- fonttools（系统字体识别）
- 纯前端 HTML/CSS/JS，无需 npm

---

## 贡献

欢迎提 Issue 和 Pull Request。

改进方向建议：
- 更多公文格式预设
- 支持页眉页脚自定义
- 表格格式处理
- 命令行批量处理模式

---

## 许可

MIT License
