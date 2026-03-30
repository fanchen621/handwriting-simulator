# 📝 教案手写仿真系统 v2.1

将电子教案转换为逼真的手写教案图片，支持双面打印，以假乱真。

## ✨ 特性

- **双面教案支持** — 正面（课题/目标/重难点/教学过程）+ 背面（板书设计/课后反思）
- **智能裁剪** — 上传照片自动检测纸张边缘
- **高真实度渲染** — 每字符随机抖动、墨水深浅变化、纸张纹理、双红线边距
- **标准教案格式** — 完整适配中国小学教案排版
- **Web 界面** — 浏览器打开即用，无需安装

## 🚀 使用

### Web 界面（推荐）
浏览器打开 `index.html` 即可使用。

### Python 命令行
```bash
# 生成演示
python render_engine.py

# 批量处理
python render_engine.py batch /path/to/lesson_plans/

# 单文件处理
python render_engine.py docx /path/to/lesson.docx
```

## 📁 文件结构

```
├── index.html          # Web 图形界面
├── render_engine.py    # Python 渲染引擎
├── demo_output/        # 示例输出
└── README.md
```

## 🖨️ 打印建议

- A4 纸，200 DPI 效果最佳
- 双面打印选"短边翻转"
- 推荐 80g+ 纸张

## ⚙️ 依赖

- Python 3.8+
- Pillow, numpy（自动安装）

## License

MIT
