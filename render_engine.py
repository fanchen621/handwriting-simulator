#!/usr/bin/env python3
"""
教案手写仿真系统 v2.1 - 高真实度渲染引擎
支持双面教案、自动裁剪、真实手写渲染
优化版：基于真实教案样张调参，逼真度最大化
"""
import os
import sys
import re
import glob
import json
import random
import math
from pathlib import Path

try:
    from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageEnhance
    import numpy as np
except ImportError:
    os.system(f"{sys.executable} -m pip install Pillow numpy")
    from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageEnhance
    import numpy as np


# ===================== CONFIG (based on real notebook analysis) =====================
class Config:
    """A4 notebook settings calibrated from real handwritten samples"""
    DPI = 200  # Higher DPI for more realistic output
    PAGE_W = int(8.27 * DPI)   # 1654px
    PAGE_H = int(11.69 * DPI)  # 2338px

    # Layout (calibrated from sample photos)
    RED_LINE_X = int(0.38 * DPI)       # Red margin ~7.6mm from left
    MARGIN_TOP = int(0.45 * DPI)
    MARGIN_RIGHT = int(0.25 * DPI)
    MARGIN_BOTTOM = int(0.3 * DPI)

    # Line spacing (tighter, matching real notebooks ~7mm)
    LINE_SPACING = int(0.28 * DPI)  # Was 0.34, now tighter like real notebooks
    NUM_LINES = (PAGE_H - MARGIN_TOP - MARGIN_BOTTOM) // LINE_SPACING

    # Text area
    TEXT_X = RED_LINE_X + int(0.18 * DPI)  # Text starts ~3.6mm after red line
    TEXT_MAX_W = PAGE_W - TEXT_X - MARGIN_RIGHT

    # Colors (calibrated from sample photos)
    PAPER_BASE = (252, 251, 245)        # Warm off-white
    LINE_COLOR = (165, 190, 215)         # Light blue-gray ruled lines
    RED_LINE = (200, 45, 45)             # Red margin
    RED_LINE_LIGHT = (200, 45, 45, 90)   # Secondary red line
    INK_BLUE = (15, 45, 140)             # Primary blue ink (rollerball pen)
    INK_BLUE_LIGHT = (30, 65, 155)       # Lighter ink variation
    INK_BLUE_DARK = (8, 25, 100)         # Darker pressure spots
    INK_FADED = (50, 80, 160, 160)       # Faded/quick strokes

    # Handwriting jitter (calibrated for natural feel)
    JITTER_X = 1.2
    JITTER_Y = 0.8
    JITTER_ROTATION = 0.015
    LINE_TILT = 0.0015                   # Very slight upward drift
    CHAR_SPREAD = 0.08                   # Character width variation ±8%

    # Section box
    BOX_PAD = 12
    BOX_RADIUS = 4
    BOX_COLOR = (15, 45, 140)
    BOX_WIDTH = 1.5


# ===================== FONT LOADER =====================
def load_font(size):
    """Load best available CJK font"""
    candidates = [
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/simkai.ttf",
        "C:/Windows/Fonts/kaiu.ttf",
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
    ]
    for fp in candidates:
        if fp and os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                continue
    return ImageFont.load_default()


# ===================== HANDWRITING ENGINE =====================
class Handwriter:
    """Renders text with realistic handwriting effects"""

    def __init__(self, font_size=24):
        self.fs = font_size
        self.font = load_font(font_size)
        self.font_b = load_font(font_size + 2)
        self.font_s = load_font(font_size - 2)  # Small font for notes

    def ink_color(self):
        """Random ink color variation"""
        r = random.random()
        if r < 0.60:
            return Config.INK_BLUE
        elif r < 0.80:
            return Config.INK_BLUE_LIGHT
        elif r < 0.92:
            return Config.INK_BLUE_DARK
        else:
            return Config.INK_FADED

    def char_width(self, ch, font=None):
        f = font or self.font
        return f.getlength(ch)

    def text_width(self, text, font=None):
        f = font or self.font
        return f.getlength(text)

    def draw_char(self, draw, ch, x, y, font=None, color=None, jitter=True):
        """Draw single character with natural variation"""
        f = font or self.font
        c = color or self.ink_color()

        if jitter and ch.strip():
            jx = x + random.gauss(0, Config.JITTER_X)
            jy = y + random.gauss(0, Config.JITTER_Y)
        else:
            jx, jy = x, y

        draw.text((jx, jy), ch, font=f, fill=c)

    def draw_line(self, draw, text, x, y, max_w=None, font=None, color=None, indent=0):
        """Draw a line of text with handwriting effect. Returns (width, actual_y)."""
        f = font or self.font
        cx = x + indent
        line_start_x = cx

        for ch in text:
            if ch == ' ':
                cx += self.fs * 0.35
                continue

            cw = self.char_width(ch, f)
            stretch = 1.0 - Config.CHAR_SPREAD/2 + random.random() * Config.CHAR_SPREAD

            # Per-character vertical jitter
            cy = y + random.gauss(0, Config.JITTER_Y * 0.4)

            c = color or self.ink_color()
            draw.text((cx + random.gauss(0, Config.JITTER_X * 0.3), cy),
                      ch, font=f, fill=c)

            cx += cw * stretch

            if max_w and (cx - line_start_x) > max_w:
                break

        return cx - line_start_x

    def wrap(self, text, max_w, font=None):
        """Word-wrap text to fit width"""
        f = font or self.font
        result = []
        for para in text.split('\n'):
            if not para.strip():
                result.append('')
                continue
            line = ''
            for ch in para:
                test = line + ch
                if f.getlength(test) > max_w and line:
                    result.append(line)
                    line = ch
                else:
                    line = test
            if line:
                result.append(line)
        return result if result else ['']


# ===================== PAGE BUILDER =====================
class PageBuilder:
    """Builds a realistic notebook page"""

    def __init__(self, hw):
        self.hw = hw

    def new_page(self):
        """Create blank ruled page"""
        img = Image.new('RGB', (Config.PAGE_W, Config.PAGE_H), Config.PAPER_BASE)
        draw = ImageDraw.Draw(img)
        self._draw_lines(draw)
        self._draw_red_margin(draw)
        return img, draw

    def _draw_lines(self, draw):
        """Draw horizontal ruled lines"""
        for i in range(Config.NUM_LINES):
            ly = Config.MARGIN_TOP + i * Config.LINE_SPACING
            # Main line
            draw.line(
                [(Config.RED_LINE_X - int(0.08 * Config.DPI), ly),
                 (Config.PAGE_W - Config.MARGIN_RIGHT, ly)],
                fill=Config.LINE_COLOR, width=1
            )

    def _draw_red_margin(self, draw):
        """Draw double red margin line (like real notebooks)"""
        x = Config.RED_LINE_X
        draw.line([(x, 8), (x, Config.PAGE_H - 8)],
                  fill=Config.RED_LINE, width=2)
        draw.line([(x + 3, 8), (x + 3, Config.PAGE_H - 8)],
                  fill=Config.RED_LINE, width=1)

    def add_texture(self, img):
        """Add paper grain and aging"""
        arr = np.array(img).astype(np.float32)

        # Paper fiber noise
        noise = np.random.normal(0, 2.5, arr.shape)
        arr += noise

        # Slight warm gradient (top lighter, bottom slightly darker)
        grad = np.linspace(0, -1.5, Config.PAGE_H).reshape(-1, 1, 1)
        grad = np.broadcast_to(grad, arr.shape)
        arr += grad

        # Very subtle color temperature variation
        arr[:, :, 0] += np.random.normal(0, 0.5, (Config.PAGE_H, Config.PAGE_W))
        arr[:, :, 2] -= np.random.normal(0, 0.3, (Config.PAGE_H, Config.PAGE_W))

        arr = np.clip(arr, 0, 255).astype(np.uint8)
        result = Image.fromarray(arr)

        # Left edge shadow (binding)
        shadow = Image.new('RGBA', result.size, (0, 0, 0, 0))
        sd = ImageDraw.Draw(shadow)
        for i in range(25):
            alpha = int(18 * (1 - i / 25))
            sd.line([(i, 0), (i, Config.PAGE_H)], fill=(0, 0, 0, alpha))
        result = result.convert('RGBA')
        result = Image.alpha_composite(result, shadow)
        return result.convert('RGB')


# ===================== SECTION RENDERER =====================
def render_box(draw, hw, title, lines, x, y, max_w, line_sp):
    """Render a bordered section (板书设计 / 课后反思)"""
    pad = Config.BOX_PAD
    inner_w = max_w - pad * 2

    # Wrap content
    wrapped = []
    for line in lines:
        wrapped.extend(hw.wrap(line, inner_w - 10))

    # Calculate box dimensions
    title_h = hw.fs + 4
    content_h = len(wrapped) * int(line_sp * 0.88)
    box_h = title_h + content_h + pad * 2 + 6
    box_w = max_w + pad

    bx = int(x - pad // 2)
    by = int(y - hw.fs - pad // 2)

    # Draw box
    draw.rounded_rectangle(
        [bx, by, bx + int(box_w), by + int(box_h)],
        radius=int(Config.BOX_RADIUS),
        outline=Config.BOX_COLOR,
        width=int(Config.BOX_WIDTH)
    )

    # Centered title
    tw = hw.text_width(title, hw.font_b)
    tx = int(bx + (box_w - tw) / 2)
    draw.text((tx, by + pad // 2 + 1), title,
              font=hw.font_b, fill=Config.INK_BLUE)

    # Content
    cy = by + title_h + pad + 4
    for line in wrapped:
        if not line.strip():
            cy += int(line_sp * 0.4)
            continue
        is_diagram = '→' in line or '｜' in line or line.startswith('看到')
        indent = 0 if is_diagram else 8
        font = hw.font if not is_diagram else hw.font
        color = Config.INK_BLUE_DARK if is_diagram else None
        hw.draw_line(draw, line, x + pad // 2 + 4, cy,
                    int(inner_w - 8), font=font, color=color, indent=indent)
        cy += int(line_sp * 0.88)

    return int(box_h)


# ===================== FRONT PAGE RENDERER =====================
def render_front(lesson, hw):
    """Render front side of lesson plan (title, goals, keypoints, main content)"""
    pb = PageBuilder(hw)
    img, draw = pb.new_page()

    x = Config.TEXT_X
    y = Config.MARGIN_TOP
    max_w = Config.TEXT_MAX_W
    lsp = Config.LINE_SPACING

    title = lesson.get('title', '')
    lesson_type = lesson.get('type', '新授课')
    goals = lesson.get('goals', '')
    keypoints = lesson.get('keypoints', '')
    content = lesson.get('front_content', '')

    # ── Title row ──
    hw.draw_line(draw, f'课题  {title}', x, y, max_w, font=hw.font_b)
    y += lsp

    # ── Type + divider ──
    hw.draw_line(draw, f'课型：{lesson_type}', x, y, max_w)
    y += lsp
    draw.line([(x, y - 3), (x + max_w, y - 3)],
              fill=Config.INK_BLUE, width=1)
    y += int(lsp * 0.45)

    # ── Goals (紧凑排版) ──
    hw.draw_line(draw, '教学目标', x, y, max_w, font=hw.font_b)
    y += lsp
    for line in hw.wrap(goals, max_w - 15):
        remaining = Config.PAGE_H - Config.MARGIN_BOTTOM - 40
        if y > remaining:
            hw.draw_line(draw, '（续背面）', x, y, max_w, font=hw.font_s)
            break
        hw.draw_line(draw, line, x + 12, y, max_w - 15)
        y += int(lsp * 0.88)

    y += int(lsp * 0.35)

    # ── Keypoints ──
    hw.draw_line(draw, '教学重难点', x, y, max_w, font=hw.font_b)
    y += lsp
    for line in hw.wrap(keypoints, max_w - 15):
        remaining = Config.PAGE_H - Config.MARGIN_BOTTOM - 40
        if y > remaining:
            hw.draw_line(draw, '（续背面）', x, y, max_w, font=hw.font_s)
            break
        hw.draw_line(draw, line, x + 12, y, max_w - 15)
        y += int(lsp * 0.88)

    y += int(lsp * 0.35)

    # ── Main teaching process ──
    hw.draw_line(draw, '教学过程', x, y, max_w, font=hw.font_b)
    y += lsp

    for raw_line in content.split('\n'):
        stripped = raw_line.strip()
        if not stripped:
            y += int(lsp * 0.3)
            continue

        remaining = Config.PAGE_H - Config.MARGIN_BOTTOM - 30
        if y > remaining:
            hw.draw_line(draw, '（续背面）', x, y, max_w, font=hw.font_s)
            break

        # Detect heading lines
        is_heading = bool(re.match(r'^[一二三四五六七八九十]+[、，]', stripped))

        wrapped = hw.wrap(stripped, max_w - (15 if not is_heading else 0))
        for wl in wrapped:
            if is_heading:
                hw.draw_line(draw, wl, x, y, max_w, font=hw.font_b)
            else:
                hw.draw_line(draw, wl, x + 15, y, max_w - 15)
            y += lsp

    img = pb.add_texture(img)
    return img


# ===================== BACK PAGE RENDERER =====================
def render_back(lesson, hw):
    """Render back side (continuation, board design, reflection)"""
    pb = PageBuilder(hw)
    img, draw = pb.new_page()

    x = Config.TEXT_X
    y = Config.MARGIN_TOP
    max_w = Config.TEXT_MAX_W
    lsp = Config.LINE_SPACING

    content = lesson.get('back_content', '')

    # ── Continuation marker ──
    hw.draw_line(draw, '（续前页）', x, y, max_w, font=hw.font_b)
    y += int(lsp * 1.3)

    lines = content.split('\n')
    in_section = False
    section_title = ''
    section_lines = []
    section_y = 0

    for raw_line in lines:
        stripped = raw_line.strip()
        if not stripped:
            if in_section:
                section_lines.append('')
            else:
                y += int(lsp * 0.3)
            continue

        remaining = Config.PAGE_H - Config.MARGIN_BOTTOM - 30
        if y > remaining:
            if in_section and section_lines:
                render_box(draw, hw, section_title, section_lines,
                          x, section_y, max_w, lsp)
            hw.draw_line(draw, '（完）', x, y, max_w, font=hw.font_s)
            break

        # Detect section headers
        is_section = False
        for kw in ['板书设计', '课后反思', '教学反思', '作业设计', '板书']:
            if kw in stripped and len(stripped) < len(kw) + 5:
                is_section = True
                # Flush previous section
                if in_section and section_lines:
                    bh = render_box(draw, hw, section_title, section_lines,
                                   x, section_y, max_w, lsp)
                    y = section_y + bh + int(lsp * 1.2)
                in_section = True
                section_title = kw
                section_lines = []
                section_y = y + hw.fs + 6
                break

        if is_section:
            continue

        if in_section:
            section_lines.append(stripped)
        else:
            is_heading = bool(re.match(r'^[一二三四五六七八九十]+[、，]', stripped))
            wrapped = hw.wrap(stripped, max_w - (15 if not is_heading else 0))
            for wl in wrapped:
                if is_heading:
                    hw.draw_line(draw, wl, x, y, max_w, font=hw.font_b)
                else:
                    hw.draw_line(draw, wl, x + 15, y, max_w - 15)
                y += lsp

    # Render final section
    if in_section and section_lines:
        render_box(draw, hw, section_title, section_lines,
                  x, section_y, max_w, lsp)

    img = pb.add_texture(img)
    return img


# ===================== DOCX / DOC READER =====================
def read_lesson_file(file_path):
    """Read lesson plan from .docx or .doc file"""
    ext = os.path.splitext(file_path)[1].lower()

    result = {
        'title': '', 'type': '新授课',
        'goals': '', 'keypoints': '',
        'front_content': '', 'back_content': '',
    }

    if ext == '.docx':
        try:
            from docx import Document
            doc = Document(file_path)
            full_text = '\n'.join(p.text for p in doc.paragraphs if p.text.strip())
            return _parse_lesson_text(full_text, result)
        except Exception as e:
            print(f"DOCX read error: {e}")

    # .doc fallback: extract text from binary
    try:
        with open(file_path, 'rb') as f:
            data = f.read()

        best_text = ''
        for enc in ['utf-16-le', 'gb18030', 'gbk', 'gb2312', 'utf-8']:
            try:
                text = data.decode(enc, errors='ignore')
                # Keep only meaningful Chinese text segments
                segments = re.findall(
                    r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef\w\s，。、；：！？\u201c\u201d\u2018\u2019（）\d.\-\(\)]+',
                    text
                )
                clean = [s.strip() for s in segments if len(s.strip()) > 3]
                joined = '\n'.join(clean)
                if len(joined) > len(best_text):
                    best_text = joined
            except Exception:
                continue

        if best_text:
            return _parse_lesson_text(best_text, result)
    except Exception as e:
        print(f"File read error: {e}")

    return result


def _parse_lesson_text(text, result):
    """Parse lesson plan text into structured data"""
    lines = text.split('\n')
    current = 'front_content'

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Detect title (look for 课题 keyword)
        m = re.search(r'课题[：:\s]*([^\r\n]{2,20})', line)
        if m and not result['title']:
            raw_title = m.group(1).strip()
            # Clean title: keep only the lesson name part
            raw_title = re.split(r'[板块一二三四五六七八九十]', raw_title)[0].strip()
            if raw_title:
                result['title'] = raw_title
            continue

        # Detect section switches
        if re.search(r'教学目标|教学目的', line) and len(line) < 15:
            current = 'goals'
            continue
        if re.search(r'教学重[点难点]|重难点|重点难点', line) and len(line) < 20:
            current = 'keypoints'
            continue
        if '板书设计' in line or '板书' == line.strip():
            current = 'back_content'
            result[current] += line + '\n'
            continue
        if '课后反思' in line or '教学反思' in line:
            current = 'back_content'
            result[current] += line + '\n'
            continue
        if re.search(r'教学过程|教学步骤', line) and len(line) < 15:
            current = 'front_content'
            continue

        result[current] += line + '\n'

    return result


# ===================== HIGH-LEVEL API =====================
def generate(lesson_data, output_dir='output', font_path=None, font_size=24):
    """Generate double-sided lesson plan images"""
    os.makedirs(output_dir, exist_ok=True)

    # Use custom font if provided
    if font_path and os.path.exists(font_path):
        try:
            hw = Handwriter.__new__(Handwriter)
            hw.fs = font_size
            hw.font = ImageFont.truetype(font_path, font_size)
            hw.font_b = ImageFont.truetype(font_path, font_size + 2)
            hw.font_s = ImageFont.truetype(font_path, font_size - 2)
        except Exception:
            hw = Handwriter(font_size)
    else:
        hw = Handwriter(font_size)

    title = re.sub(r'[\\/:*?"<>|\r\n]', '_', lesson_data.get('title', 'untitled'))
    # Keep filename short
    title = title[:30].strip('_. ')

    print(f"  渲染正面: {title}")
    front = render_front(lesson_data, hw)
    front_path = os.path.join(output_dir, f"{title}_正面.png")
    front.save(front_path, dpi=(Config.DPI, Config.DPI))

    print(f"  渲染背面: {title}")
    back = render_back(lesson_data, hw)
    back_path = os.path.join(output_dir, f"{title}_背面.png")
    back.save(back_path, dpi=(Config.DPI, Config.DPI))

    return {'front': front_path, 'back': back_path, 'title': title}


def batch(input_dir, output_dir='output', font_size=24):
    """Batch process all .doc/.docx files"""
    files = glob.glob(os.path.join(input_dir, '*.doc')) + \
            glob.glob(os.path.join(input_dir, '*.docx'))
    files = [f for f in files if not os.path.basename(f).startswith('~')]

    if not files:
        print(f"未找到教案文件: {input_dir}")
        return []

    results = []
    for fpath in sorted(files):
        print(f"\n处理: {os.path.basename(fpath)}")
        try:
            data = read_lesson_file(fpath)
            if data.get('front_content') or data.get('back_content'):
                r = generate(data, output_dir, font_size=font_size)
                results.append(r)
            else:
                print("  跳过: 未提取到内容")
        except Exception as e:
            print(f"  错误: {e}")

    print(f"\n完成! 共生成 {len(results)} 份教案 → {output_dir}/")
    return results


# ===================== DEMO =====================
def demo():
    """Generate demo based on real sample lesson plan"""
    lesson = {
        'title': '3 天窗',
        'type': '新授课',
        'goals': (
            '1.认识"慰、藉"等3个生字，会写"慰、藉、瞥"等10个生字，'
            '理解"慰藉、威力"等词语。\n'
            '2.正确、流利、有感情地朗读课文，理解课文内容，'
            '体会"小小的天窗是你唯一的慰藉"这句话的含义。\n'
            '3.学习作者运用丰富的想象，把看到的事物写得更具体、更生动的写作方法。'
        ),
        'keypoints': (
            '重点：理解课文内容，了解天窗给乡下孩子带来的无尽遐想和快乐。\n'
            '难点：理解"小小的天窗是你唯一的慰藉"的含义，'
            '体会作者对天窗的特殊感情。'
        ),
        'front_content': (
            '一、激趣导入，揭示课题\n'
            '1.同学们，你们见过天窗吗？谁知道什么是天窗？\n'
            '2.出示天窗图片，简介天窗：天窗是在屋顶开的一个小窗子，用来采光和通风。\n'
            '3.板书课题，齐读课题。\n'
            '4.质疑：读了课题，你想知道什么？\n'
            '二、初读课文，整体感知\n'
            '1.自由朗读课文，读准字音，读通句子。\n'
            '2.检查生字词认读情况。\n'
            '3.默读课文，思考：什么是天窗？天窗在什么时候成为孩子们"唯一的慰藉"？\n'
            '三、精读课文，品味语言\n'
            '1.学习第1-3自然段\n'
            '（1）指名读第1-3自然段。\n'
            '（2）说说天窗是什么样子的？为什么要开天窗？\n'
            '（3）交流：天窗的位置、样子和作用。\n'
            '2.学习第4-7自然段\n'
            '（1）默读第4-7自然段，画出表现孩子们心情变化的词句。\n'
            '（2）交流：下雨天，孩子们被关在屋里时的心情怎样？\n'
            '透过天窗看到了什么？想到了什么？\n'
            '（3）品读重点句子，体会想象的丰富。\n'
            '（4）指导朗读，读出孩子们的快乐。\n'
            '四、课堂小结\n'
            '天窗虽然很小，却给孩子们带来了无穷的快乐和想象。'
        ),
        'back_content': (
            '五、学习第8自然段\n'
            '1.齐读最后一个自然段。\n'
            '2.理解"发明这天窗的大人们，是应得感谢的"这句话的含义。\n'
            '3.小结：天窗不仅给了孩子们一个看世界的窗口，更给了他们想象的翅膀。\n'
            '六、总结全文，拓展延伸\n'
            '1.用自己的话说说课文的主要内容。\n'
            '2.体会作者对天窗的感情，感受天窗给孩子们带来的快乐。\n'
            '3.拓展：你有没有类似的经历？在无聊或孤独的时候，是什么给了你快乐和想象？\n'
            '板书设计\n'
            '天窗\n'
            '屋子→天窗→想象\n'
            '木板窗→天窗→慰藉\n'
            '看到的→想到的\n'
            '一颗星→无数星\n'
            '一朵云→无数云\n'
            '一条黑影→蝙蝠、夜莺……\n'
            '课后反思\n'
            '本课教学中，我注重引导学生抓住关键词句理解课文内容，'
            '通过品读、想象、朗读等多种方式感受天窗给孩子们带来的快乐。'
            '在教学"想象"这部分内容时，我引导学生展开想象的翅膀，'
            '用自己的语言描述透过天窗看到的和想到的内容，'
            '激发了学生的学习兴趣。教学效果较好，学生参与度高。'
            '不足之处是部分学生对"唯一的慰藉"理解还不够深入，'
            '需要在第二课时进一步引导体会。'
        ),
    }
    return generate(lesson, 'demo_output')


if __name__ == '__main__':
    if len(sys.argv) > 1:
        if sys.argv[1] == 'batch' and len(sys.argv) > 2:
            batch(sys.argv[2])
        elif sys.argv[1] == 'docx' and len(sys.argv) > 2:
            data = read_lesson_file(sys.argv[2])
            generate(data)
        elif sys.argv[1] == 'doc' and len(sys.argv) > 2:
            data = read_lesson_file(sys.argv[2])
            print(json.dumps({k: v[:100]+'...' if len(v)>100 else v
                              for k, v in data.items()}, ensure_ascii=False, indent=2))
        else:
            print("用法:")
            print("  python render_engine.py              # 生成演示")
            print("  python render_engine.py batch <目录>  # 批量处理")
            print("  python render_engine.py docx <文件>   # 单文件处理")
            print("  python render_engine.py doc <文件>    # 预览提取内容")
    else:
        demo()
