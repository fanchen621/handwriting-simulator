#!/usr/bin/env python3
"""
教案手写仿真系统 v4.0 - 改进版渲染引擎
改进：纯白纸无横线、无方框结构块、更自然的手写效果、内容完整性保障
"""
import os
import sys
import re
import glob
import json
import random
import math
import argparse
from pathlib import Path

try:
    from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageEnhance
    import numpy as np
except ImportError:
    os.system(f"{sys.executable} -m pip install Pillow numpy")
    from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageEnhance
    import numpy as np

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


# ===================== CONFIG =====================
class Config:
    """A4 settings - plain paper, no lines, like real scanned documents"""
    DPI = 200
    PAGE_W = int(8.27 * DPI)   # 1654px
    PAGE_H = int(11.69 * DPI)  # 2338px

    # Layout - slightly wider margins for natural look
    RED_LINE_X = int(0.35 * DPI)
    MARGIN_TOP = int(0.40 * DPI)
    MARGIN_RIGHT = int(0.30 * DPI)
    MARGIN_BOTTOM = int(0.25 * DPI)

    LINE_SPACING = int(0.28 * DPI)
    NUM_LINES = (PAGE_H - MARGIN_TOP - MARGIN_BOTTOM) // LINE_SPACING

    TEXT_X = RED_LINE_X + int(0.15 * DPI)
    TEXT_MAX_W = PAGE_W - TEXT_X - MARGIN_RIGHT

    # Colors
    PAPER_WHITE = (255, 255, 255)
    PAPER_CREAM = (252, 250, 243)
    RED_LINE = (200, 50, 50)
    INK_BLUE = (15, 45, 140)
    INK_BLUE_LIGHT = (30, 65, 155)
    INK_BLACK = (17, 17, 17)
    INK_RED = (200, 45, 45)

    # Handwriting jitter (base) - increased for more natural look
    JITTER_X = 1.8
    JITTER_Y = 1.2
    CHAR_SPREAD = 0.10


# ===================== HANDWRITING STYLE =====================
class HandwritingStyle:
    def __init__(self):
        self.font_size = 22
        self.line_spacing_mult = 1.0
        self.position_disorder = 0.35
        self.stroke_disorder = 0.25
        self.scribble_prob = 0.05
        self.ink_variation = 2
        self.ink_color = 'blue'
        self.paper_type = 'plain'  # Default to plain white paper
        self.font_name = 'kai'
        self.custom_font_path = None
        self.custom_paper_path = None

    @property
    def ink_rgb(self):
        colors = {
            'blue': Config.INK_BLUE,
            'blue2': Config.INK_BLUE_LIGHT,
            'black': Config.INK_BLACK,
            'red': Config.INK_RED,
        }
        return colors.get(self.ink_color, Config.INK_BLUE)


# ===================== FONT LOADER =====================
def load_font(size, font_name='kai', custom_path=None):
    if custom_path and os.path.exists(custom_path):
        try:
            return ImageFont.truetype(custom_path, size)
        except Exception:
            pass

    font_map = {
        'kai': [
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
            "C:/Windows/Fonts/simkai.ttf",
            "C:/Windows/Fonts/kaiu.ttf",
            "/System/Library/Fonts/STHeiti Light.ttc",
        ],
        'fang': [
            "C:/Windows/Fonts/simfang.ttf",
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        ],
        'hei': [
            "C:/Windows/Fonts/simhei.ttf",
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        ],
    }

    candidates = font_map.get(font_name, font_map['kai'])
    candidates += [
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "/System/Library/Fonts/PingFang.ttc",
    ]

    for fp in candidates:
        if os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                continue
    return ImageFont.load_default()


# ===================== HANDWRITING ENGINE =====================
class Handwriter:
    def __init__(self, style: HandwritingStyle = None):
        self.style = style or HandwritingStyle()
        s = self.style
        self.fs = s.font_size
        self.font = load_font(s.font_size, s.font_name, s.custom_font_path)
        self.font_b = load_font(s.font_size + 2, s.font_name, s.custom_font_path)
        self.font_s = load_font(s.font_size - 2, s.font_name, s.custom_font_path)

    def ink_color(self):
        base = self.style.ink_rgb
        v = self.style.ink_variation
        if v == 0:
            return base
        r = random.random()
        if v <= 1:
            return base if r < 0.85 else (
                tuple(min(255, c + 20) for c in base) if r < 0.95
                else tuple(max(0, c - 20) for c in base)
            )
        if r < 0.50:
            return base
        elif r < 0.70:
            return tuple(min(255, c + 15 + v * 5) for c in base)
        elif r < 0.85:
            return tuple(max(0, c - 15 - v * 5) for c in base)
        else:
            return tuple(min(255, c + 30) for c in base[:3])

    def char_width(self, ch, font=None):
        f = font or self.font
        return f.getlength(ch)

    def text_width(self, text, font=None):
        f = font or self.font
        return f.getlength(text)

    def draw_char(self, draw, ch, x, y, font=None, color=None):
        """Draw single character with position + stroke disorder + baseline drift"""
        f = font or self.font
        c = color or self.ink_color()
        pd = self.style.position_disorder
        sd = self.style.stroke_disorder

        if ch.strip():
            jx = x + random.gauss(0, Config.JITTER_X * (1 + pd * 4))
            jy = y + random.gauss(0, Config.JITTER_Y * (1 + pd * 3))
            # Baseline drift - subtle wave
            jy += math.sin(x * 0.02) * pd * 2
        else:
            jx, jy = x, y

        # Stroke disorder: draw ghost/offset copy
        if sd > 0 and ch.strip() and random.random() < sd * 0.2:
            ghost_c = tuple(list(c[:3]) + [35])
            draw.text(
                (jx + random.gauss(0, sd * 3), jy + random.gauss(0, sd * 2)),
                ch, font=f, fill=ghost_c
            )

        # Ink density variation
        if pd > 0.1 and random.random() < 0.15:
            alpha_mod = int(random.randint(180, 255))
            c = tuple(list(c[:3]) + [alpha_mod]) if len(c) == 3 else c

        draw.text((jx, jy), ch, font=f, fill=c)

    def draw_line(self, draw, text, x, y, max_w=None, font=None, color=None, indent=0):
        f = font or self.font
        cx = x + indent
        pd = self.style.position_disorder
        sd = self.style.stroke_disorder
        base_phase = random.random() * math.pi * 2

        for i, ch in enumerate(text):
            if ch == ' ':
                cx += self.fs * 0.35
                continue
            cw = self.char_width(ch, f)
            stretch = 1.0 - Config.CHAR_SPREAD / 2 + random.random() * Config.CHAR_SPREAD
            cy = y + random.gauss(0, Config.JITTER_Y * (0.4 + pd * 2))
            # Baseline drift
            cy += math.sin(base_phase + i * 0.3) * pd * 1.5
            c = color or self.ink_color()

            # Stroke disorder ghost
            if sd > 0 and ch.strip() and random.random() < sd * 0.2:
                draw.text(
                    (cx + random.gauss(0, sd * 2.5), cy + random.gauss(0, sd * 1.5)),
                    ch, font=f, fill=(*c[:3], 30)
                )

            draw.text(
                (cx + random.gauss(0, Config.JITTER_X * (0.3 + pd * 1.5)), cy),
                ch, font=f, fill=c
            )
            cx += cw * stretch

            if max_w and (cx - x - indent) > max_w:
                break

        return cx - x - indent

    def wrap(self, text, max_w, font=None):
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


# ===================== PAPER BUILDER =====================
class PaperBuilder:
    """Plain paper - no horizontal lines, no grid. Like real scanned documents."""

    @staticmethod
    def create_page(style: HandwritingStyle):
        if style.custom_paper_path and os.path.exists(style.custom_paper_path):
            try:
                img = Image.open(style.custom_paper_path).convert('RGB')
                img = img.resize((Config.PAGE_W, Config.PAGE_H), Image.LANCZOS)
                return img
            except Exception:
                pass

        if style.paper_type == 'cream':
            bg = Config.PAPER_CREAM
        else:
            bg = Config.PAPER_WHITE  # Default: plain white paper

        img = Image.new('RGB', (Config.PAGE_W, Config.PAGE_H), bg)
        draw = ImageDraw.Draw(img)

        # NO horizontal lines drawn - plain paper only
        # Subtle red margin line only
        PaperBuilder._draw_red_margin(draw)
        return img

    @staticmethod
    def _draw_red_margin(draw):
        x = Config.RED_LINE_X
        draw.line([(x, 10), (x, Config.PAGE_H - 10)],
                  fill=Config.RED_LINE, width=1)

    @staticmethod
    def add_texture(img):
        """Subtle paper grain and edge shadow - like real paper"""
        arr = np.array(img).astype(np.float32)
        # Very subtle paper fiber noise
        noise = np.random.normal(0, 1.8, arr.shape)
        arr += noise
        arr = np.clip(arr, 0, 255).astype(np.uint8)
        result = Image.fromarray(arr)

        # Left edge shadow (binding)
        shadow = Image.new('RGBA', result.size, (0, 0, 0, 0))
        sd = ImageDraw.Draw(shadow)
        for i in range(20):
            alpha = int(15 * (1 - i / 20))
            sd.line([(i, 0), (i, Config.PAGE_H)], fill=(0, 0, 0, alpha))
        result = result.convert('RGBA')
        result = Image.alpha_composite(result, shadow)
        return result.convert('RGB')


# ===================== SCRIBBLE ENGINE =====================
def add_scribbles(draw, hw, prob, ink_color):
    if prob <= 0:
        return
    count = int(prob * 12)
    W, H = Config.PAGE_W, Config.PAGE_H
    mx = Config.RED_LINE_X

    for _ in range(count):
        sx = mx + 30 + random.random() * (W - mx - 60)
        sy = Config.MARGIN_TOP + random.random() * (H - Config.MARGIN_TOP - 40)
        stype = random.randint(0, 2)
        alpha = random.randint(50, 130)
        c = (*ink_color[:3], alpha)

        if stype == 0:
            r = 4 + random.random() * 10
            pts = []
            for a in [i * 0.4 for i in range(int(math.pi * 2 / 0.4) + 1)]:
                rx = sx + math.cos(a) * r + (random.random() - 0.5) * 3
                ry = sy + math.sin(a) * r + (random.random() - 0.5) * 3
                pts.append((rx, ry))
            if len(pts) > 2:
                draw.line(pts, fill=c, width=1)
        elif stype == 1:
            draw.line([(sx - 6, sy - 3), (sx + 6, sy + 3)], fill=c, width=1)
            draw.line([(sx - 6, sy + 3), (sx + 6, sy - 3)], fill=c, width=1)
        else:
            pts = []
            for dx in range(-12, 13, 2):
                wx = sx + dx
                wy = sy + math.sin(dx * 0.5) * 1.5
                pts.append((wx, wy))
            if len(pts) > 1:
                draw.line(pts, fill=c, width=1)


# ===================== FRONT PAGE RENDERER =====================
def render_front(lesson, hw):
    pb = PaperBuilder()
    img = pb.create_page(hw.style)
    draw = ImageDraw.Draw(img)

    x = Config.TEXT_X
    y = Config.MARGIN_TOP
    max_w = Config.TEXT_MAX_W
    lsp = int(Config.LINE_SPACING * hw.style.line_spacing_mult)

    title = lesson.get('title', '')
    lesson_type = lesson.get('type', '新授课')
    goals = lesson.get('goals', '')
    keypoints = lesson.get('keypoints', '')
    content = lesson.get('front_content', '')

    # Title
    hw.draw_line(draw, f'课题  {title}', x, y, max_w, font=hw.font_b)
    y += lsp
    hw.draw_line(draw, f'课型：{lesson_type}', x, y, max_w)
    y += lsp
    # Subtle separator
    draw.line([(x, y - 2), (x + max_w, y - 2)], fill=(*hw.style.ink_rgb[:3], 100), width=1)
    y += int(lsp * 0.4)

    # Goals
    hw.draw_line(draw, '教学目标', x, y, max_w, font=hw.font_b)
    y += lsp
    for line in hw.wrap(goals, max_w - 12):
        if y > Config.PAGE_H - Config.MARGIN_BOTTOM - 30:
            hw.draw_line(draw, '（续背面）', x, y, max_w, font=hw.font_s)
            break
        hw.draw_line(draw, line, x + 8, y, max_w - 12)
        y += int(lsp * 0.85)
    y += int(lsp * 0.3)

    # Keypoints
    hw.draw_line(draw, '教学重难点', x, y, max_w, font=hw.font_b)
    y += lsp
    for line in hw.wrap(keypoints, max_w - 12):
        if y > Config.PAGE_H - Config.MARGIN_BOTTOM - 30:
            hw.draw_line(draw, '（续背面）', x, y, max_w, font=hw.font_s)
            break
        hw.draw_line(draw, line, x + 8, y, max_w - 12)
        y += int(lsp * 0.85)
    y += int(lsp * 0.3)

    # Teaching process
    hw.draw_line(draw, '教学过程', x, y, max_w, font=hw.font_b)
    y += lsp
    for raw_line in content.split('\n'):
        stripped = raw_line.strip()
        if not stripped:
            y += int(lsp * 0.2)
            continue
        if y > Config.PAGE_H - Config.MARGIN_BOTTOM - 25:
            hw.draw_line(draw, '（续背面）', x, y, max_w, font=hw.font_s)
            break
        is_heading = bool(re.match(r'^[一二三四五六七八九十]+[、，,.]', stripped))
        is_sub = bool(re.match(r'^（?\d+[）)]', stripped))
        indent = 0 if is_heading else (12 if is_sub else 8)
        wrapped = hw.wrap(stripped, max_w - indent)
        for wl in wrapped:
            if is_heading:
                hw.draw_line(draw, wl, x + indent, y, max_w, font=hw.font_b)
            else:
                hw.draw_line(draw, wl, x + indent, y, max_w - indent)
            y += lsp

    # Scribbles
    add_scribbles(draw, hw, hw.style.scribble_prob, hw.style.ink_rgb)

    img = pb.add_texture(img)
    return img


# ===================== BACK PAGE RENDERER =====================
def render_back(lesson, hw):
    pb = PaperBuilder()
    img = pb.create_page(hw.style)
    draw = ImageDraw.Draw(img)

    x = Config.TEXT_X
    y = Config.MARGIN_TOP
    max_w = Config.TEXT_MAX_W
    lsp = int(Config.LINE_SPACING * hw.style.line_spacing_mult)

    content = lesson.get('back_content', '')

    hw.draw_line(draw, '（续前页）', x, y, max_w, font=hw.font_b)
    y += int(lsp * 1.2)

    lines = content.split('\n')
    for raw_line in lines:
        stripped = raw_line.strip()
        if not stripped:
            y += int(lsp * 0.2)
            continue

        if y > Config.PAGE_H - Config.MARGIN_BOTTOM - 25:
            hw.draw_line(draw, '（完）', x, y, max_w, font=hw.font_s)
            break

        # Section headers - bold but NO box, just like real handwritten docs
        is_section = False
        for kw in ['板书设计', '课后反思', '教学反思', '板书', '教后反思', '作业设计']:
            if kw in stripped and len(stripped) < len(kw) + 5:
                is_section = True
                y += int(lsp * 0.3)
                hw.draw_line(draw, stripped, x, y, max_w, font=hw.font_b)
                y += lsp
                break
        if is_section:
            continue

        is_heading = bool(re.match(r'^[五六七八九十][、，,.]', stripped))
        wrapped = hw.wrap(stripped, max_w - (0 if is_heading else 8))
        for wl in wrapped:
            if is_heading:
                hw.draw_line(draw, wl, x, y, max_w, font=hw.font_b)
            else:
                hw.draw_line(draw, wl, x + 8, y, max_w - 8)
            y += lsp

    # Scribbles
    add_scribbles(draw, hw, hw.style.scribble_prob, hw.style.ink_rgb)

    img = pb.add_texture(img)
    return img


# ===================== DOCX READER (ROBUST) =====================
def read_lesson_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    result = {
        'title': '', 'type': '新授课',
        'goals': '', 'keypoints': '',
        'front_content': '', 'back_content': '',
    }

    if ext == '.docx' and HAS_DOCX:
        try:
            doc = Document(file_path)
            full_text = '\n'.join(p.text for p in doc.paragraphs if p.text.strip())
            return _parse_lesson_text(full_text, result)
        except Exception as e:
            print(f"DOCX read error: {e}")

    # Fallback: binary extraction
    try:
        with open(file_path, 'rb') as f:
            data = f.read()
        best_text = ''
        for enc in ['utf-16-le', 'gb18030', 'gbk', 'gb2312', 'utf-8']:
            try:
                text = data.decode(enc, errors='ignore')
                segments = re.findall(
                    r'[\u4e00-\u9fff\u3000-303f\uff00-\uffef\w\s，。、；：！？\u201c\u201d\u2018\u2019（）\d.\-\(\)]+',
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
    """Robust lesson plan parser - handles many Chinese document formats"""
    lines = text.split('\n')
    current = 'front_content'
    title_found = False

    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue

        # Title detection - multiple patterns
        if not title_found:
            m = re.search(r'(?:课题|题目|篇名|课文)[：:\s]*(.+)', line)
            if m:
                t = m.group(1).strip().split()[0]
                t = re.sub(r'[板块一二三四五六七八九十\d.].*$', '', t).strip()
                if 2 <= len(t) <= 20:
                    result['title'] = t
                    title_found = True
                    continue
            # "3 天窗" or "第X课 XXX" patterns
            if i < 5:
                m = re.match(r'(?:\d+[\s.、]|第[一二三四五六七八九十\d]+[课章单元])\s*(.{2,15})', line)
                if m:
                    result['title'] = m.group(1).strip()
                    title_found = True
                    continue

        # Section headers - very flexible matching
        if re.search(r'^(教学目标|教学目的|教学要求|学习目标)\s*$', line):
            current = 'goals'
            continue
        if re.search(r'^(教学重点|教学难点|教学重难点|重难点|重点难点|教学重点和难点)\s*$', line):
            current = 'keypoints'
            continue
        if re.search(r'^(教学过程|教学步骤|教学流程|教学环节|教学程序)\s*$', line):
            current = 'front_content'
            continue

        # Back content triggers
        if re.search(r'^(板书设计|板书)\s*$', line):
            current = 'back_content'
            result[current] += line + '\n'
            continue
        if re.search(r'^(课后反思|教学反思|教后反思|课后小结|教学后记)\s*$', line):
            current = 'back_content'
            result[current] += line + '\n'
            continue

        # Chinese numeral section markers indicating back page
        if re.match(r'^[五六七八九十][、，,.]', line) and current == 'front_content':
            current = 'back_content'

        result[current] += line + '\n'

    # Fallback title
    if not result['title']:
        for l in lines:
            t = l.strip()
            if 2 <= len(t) <= 15 and not re.match(r'^[一二三四五六七八九十\d]', t) \
               and not re.match(r'^[教板课重]', t):
                result['title'] = t
                break
        if not result['title']:
            result['title'] = '未命名教案'

    # Clean up
    for k in ['goals', 'keypoints', 'front_content', 'back_content']:
        result[k] = result[k].strip()

    return result


# ===================== HIGH-LEVEL API =====================
def generate(lesson_data, output_dir='output', style=None):
    os.makedirs(output_dir, exist_ok=True)
    style = style or HandwritingStyle()
    hw = Handwriter(style)

    title = re.sub(r'[\\/:*?"<>|\r\n]', '_', lesson_data.get('title', 'untitled'))
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


def batch(input_dir, output_dir='output', style=None):
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
                r = generate(data, output_dir, style=style)
                results.append(r)
            else:
                print("  跳过: 未提取到内容")
        except Exception as e:
            print(f"  错误: {e}")

    print(f"\n完成! 共生成 {len(results)} 份教案 → {output_dir}/")
    return results


def demo():
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
            '3.拓展：你有没有类似的经历？\n'
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
    parser = argparse.ArgumentParser(description='教案手写仿真系统 v4.0')
    parser.add_argument('command', nargs='?', default='demo',
                        choices=['demo', 'batch', 'docx', 'doc'],
                        help='运行模式')
    parser.add_argument('path', nargs='?', help='文件或目录路径')
    parser.add_argument('--output', '-o', default='output', help='输出目录')
    parser.add_argument('--font-size', type=int, default=22, help='字号')
    parser.add_argument('--position-disorder', type=float, default=0.35,
                        help='文字位置凌乱度 (0.0-1.0)')
    parser.add_argument('--stroke-disorder', type=float, default=0.25,
                        help='字体笔画凌乱度 (0.0-1.0)')
    parser.add_argument('--scribble', type=float, default=0.05,
                        help='随机勾画涂改概率 (0.0-0.30)')
    parser.add_argument('--ink', default='blue',
                        choices=['blue', 'blue2', 'black', 'red'],
                        help='墨水颜色')
    parser.add_argument('--paper', default='plain',
                        choices=['plain', 'cream'],
                        help='纸张类型')
    parser.add_argument('--font', default='kai',
                        help='字体 (kai/fang/hei 或自定义字体路径)')
    parser.add_argument('--custom-paper', help='自定义纸张图片路径')

    args = parser.parse_args()

    style = HandwritingStyle()
    style.font_size = args.font_size
    style.position_disorder = args.position_disorder
    style.stroke_disorder = args.stroke_disorder
    style.scribble_prob = min(args.scribble, 0.30)
    style.ink_color = args.ink
    style.paper_type = args.paper

    if os.path.exists(args.font):
        style.custom_font_path = args.font
    else:
        style.font_name = args.font

    if args.custom_paper and os.path.exists(args.custom_paper):
        style.custom_paper_path = args.custom_paper

    if args.command == 'demo':
        demo()
    elif args.command == 'batch' and args.path:
        batch(args.path, args.output, style=style)
    elif args.command == 'docx' and args.path:
        data = read_lesson_file(args.path)
        generate(data, args.output, style=style)
    elif args.command == 'doc' and args.path:
        data = read_lesson_file(args.path)
        print(json.dumps({k: v[:100] + '...' if len(v) > 100 else v
                          for k, v in data.items()}, ensure_ascii=False, indent=2))
    else:
        parser.print_help()
