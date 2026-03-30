# 📝 教案手写仿真系统 v3.0

将电子教案转换为逼真的手写教案图片，支持双面打印、批量DOCX导入、PDF导出、自定义字体/纸张。

## ✨ 特性

- **双面教案支持** — 正面（课题/目标/重难点/教学过程）+ 背面（板书设计/课后反思）
- **批量导入DOCX** — 支持导入 .docx 文件，自动解析教案结构，批量渲染生成PDF打印
- **智能裁剪** — 上传照片自动检测纸张边缘、提取布局参数
- **仿手写效果** — 仿照 autohanding.com，支持：
  - 随机勾画涂改概率（0-30%）
  - 文字位置凌乱度（0-100%）
  - 字体笔画凌乱度（0-100%）
- **多种纸张样式** — 横线本、田字格、纯白纸、米黄纸
- **自定义字体** — 上传 .ttf/.otf 字体文件，自定义命名
- **自定义纸张** — 上传纸张照片，系统复刻纸张样式
- **PDF导出** — 单份或批量导出PDF，适配打印
- **Web 界面** — 浏览器打开即用，无需安装

## 🚀 使用

### Web 界面（推荐）
浏览器打开 `index.html` 即可使用。

### Python 命令行

```bash
# 生成演示
python render_engine.py demo

# 批量处理
python render_engine.py batch /path/to/lesson_plans/

# 单文件处理
python render_engine.py docx /path/to/lesson.docx

# 自定义参数
python render_engine.py demo --position-disorder 0.3 --stroke-disorder 0.2 --scribble 0.1 --ink blue --paper lined

# 使用自定义字体
python render_engine.py demo --font /path/to/font.ttf

# 使用自定义纸张
python render_engine.py demo --custom-paper /path/to/paper.jpg
```

### 命令行参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--font-size` | 字号 | 24 |
| `--position-disorder` | 文字位置凌乱度 (0.0-1.0) | 0.0 |
| `--stroke-disorder` | 字体笔画凌乱度 (0.0-1.0) | 0.0 |
| `--scribble` | 随机勾画涂改概率 (0.0-0.30) | 0.0 |
| `--ink` | 墨水颜色 (blue/blue2/black/red) | blue |
| `--paper` | 纸张类型 (lined/grid/plain/cream) | lined |
| `--font` | 字体名称或 .ttf 路径 | kai |
| `--custom-paper` | 自定义纸张图片路径 | - |
| `--output` | 输出目录 | output |

## 📁 文件结构

```
├── index.html          # Web 图形界面（含所有功能）
├── render_engine.py    # Python 渲染引擎
├── demo_output/        # 示例输出
└── README.md
```

## 📦 依赖

### Web 端（CDN自动加载）
- [Mammoth.js](https://github.com/mwilliamson/mammoth.js) — DOCX 解析
- [jsPDF](https://github.com/parallax/jsPDF) — PDF 生成

### Python 端
- Python 3.8+
- Pillow, numpy（自动安装）
- python-docx（可选，用于 .docx 解析）

```bash
pip install Pillow numpy python-docx
```

## 🖨️ 打印建议

- A4 纸，200 DPI 效果最佳
- 双面打印选"短边翻转"
- 推荐 80g+ 纸张

## License

MIT
