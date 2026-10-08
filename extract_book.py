#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图书PDF提取脚本
=================
从图书PDF中提取文字+图片，输出 txt/ + img/ + 目录.xlsx 文件包。

硬约束：
  1. 不提取页眉、页脚、页码
  2. 表格/图表/统计图 → 截取为图片
  3. 文字精确识别（不出现乱码；图片中含文字则整张截为图片）
  4. 目录三级结构（章/节/目）

用法：
  python extract_book.py <pdf路径> [输出目录]
  python extract_book.py NH0633.pdf 柔石画传
"""

import sys
import os
import re
import pymupdf
import openpyxl
from pathlib import Path
from openpyxl.styles import Font, Alignment
from difflib import SequenceMatcher

# ============================================================
# 配置参数（基于 NH0633.pdf 探索结果）
# ============================================================
PAGE_W = 1937.0
PAGE_H = 2820.0

# 页眉页脚过滤
FOOT_Y_THRESHOLD = 2540.0   # y > 此值 → 页脚
HEAD_Y_THRESHOLD = 200.0    # y < 此值 → 页眉
RIGHT_EDGE_X = 1700.0       # x > 此值且宽度<100 → 右侧竖排装饰
RIGHT_EDGE_MAX_W = 100.0

# 全页背景图过滤
FULLPAGE_W_RATIO = 0.95
FULLPAGE_H_RATIO = 0.95

# 节标题检测
SECTION_MIN_SZ = 49.0       # 最小字号
SECTION_MAX_WIDTH = 1000.0  # 最大宽度（排除正文长句）
SECTION_MIN_LEN = 4
SECTION_MAX_LEN = 30

# 章扉页检测
CHAPTER_MIN_SZ = 55.0       # 章标题最小字号
CHAPTER_MAX_ITEMS = 6       # 章扉页最多文字block数

# 图片说明
IMG_CAPTION_MAX_SZ = 38.0   # 图片说明最大字号
IMG_CAPTION_Y_GAP = 60.0    # 图片说明与图片底部的y差距

# 小图片过滤（装饰性小图 / OCR碎片）
IMG_MIN_W = 50.0            # 图片最小宽度
IMG_MIN_H = 50.0            # 图片最小高度
IMG_MIN_AREA = 50000.0      # 图片最小面积（过滤OCR文字碎片）
IMG_MAX_RATIO = 8.0         # 最大宽高比（过滤文字行碎片）

# OCR噪声过滤（正文46-50，图片说明35）
NOISE_MIN_SZ = 28.0         # 文字最小字号（低于此视为OCR噪声）

# 图片文字页检测（页面大部分文字字号远小于正文 → 整页截为图片）
IMG_TEXT_BIG_SZ = 40.0      # 视为"正文文字"的字号阈值
IMG_TEXT_BIG_RATIO = 0.2    # 正文文字占比低于此值 → 图片文字页

# 特殊标题（前置/后置部分）
SPECIAL_TITLES = [
    '龙华英烈画传系列丛书编委会',
    '出版说明',
    '目录',
    '柔石大事年表',
    '参考文献',
    '后记',
]

# 中文数字（章标题前缀）
CN_NUMS = '一二三四五六七八九十'


# ============================================================
# 过滤检测函数
# ============================================================
def is_fullpage_image(bbox):
    """检测全页背景图"""
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]
    return w > PAGE_W * FULLPAGE_W_RATIO and h > PAGE_H * FULLPAGE_H_RATIO


def is_footer(bbox):
    """检测页脚"""
    return bbox[1] > FOOT_Y_THRESHOLD


def is_header(bbox):
    """检测页眉"""
    return bbox[3] < HEAD_Y_THRESHOLD


def is_right_edge_deco(bbox, text=''):
    """检测右侧竖排装饰文字（保留含'目录'的特殊标题）"""
    w = bbox[2] - bbox[0]
    if bbox[0] > RIGHT_EDGE_X and w < RIGHT_EDGE_MAX_W:
        if '目录' in text or '目 录' in text:
            return False
        return True
    return False


def _chinese_ratio(text):
    """计算中文字符占比"""
    if not text:
        return 0.0
    cn = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
    return cn / len(text)


def is_section_title(text, max_sz, bbox):
    """检测节标题：含'：' + 字号>49 + 宽度<1000 + 长度4-30 + 非数字开头; 中文字符占比>60%"""
    w = bbox[2] - bbox[0]
    if '：' not in text:
        return False
    if max_sz < SECTION_MIN_SZ:
        return False
    if w > SECTION_MAX_WIDTH:
        return False
    if len(text) < SECTION_MIN_LEN or len(text) > SECTION_MAX_LEN:
        return False
    if text[0].isdigit():
        return False
    # 排除含多个冒号（如"编委会主任：严爱云 副主任：王旭杰"）
    if text.count('：') > 1:
        return False
    # 中文字符占比>60%（排除OCR噪声）
    if _chinese_ratio(text) < 0.6:
        return False
    # 排除以句号/感叹号结尾的正文句子（不排除引号——节标题常含引号）
    if text[-1] in '。！':
        return False
    # 冒号前后部分
    parts = text.split('：', 1)
    prefix = parts[0]
    suffix = parts[1] if len(parts) > 1 else ''
    if len(prefix) < 1 or len(prefix) > 8:
        return False
    # 排除前缀含书名号》但不含《的（如"后记》："是正文引用，但"《疯人》"是节标题）
    if '》' in prefix and '《' not in prefix:
        return False
    if '】' in prefix and '【' not in prefix:
        return False
    # 排除含"编委"的（编委会页内容）
    if '编委' in text:
        return False
    # 冒号后内容含句号→大概率是正文句子（节标题通常不含句号）
    if '。' in suffix:
        return False
    return True


def is_special_title(text, max_sz, bbox):
    """检测特殊标题（编委会、出版说明、目录、大事年表等）"""
    if max_sz < 45:
        return False
    w = bbox[2] - bbox[0]
    if w > 800:
        return False
    # 去除空格后模糊匹配
    text_compact = text.replace(' ', '').replace('\u3000', '')
    for st in SPECIAL_TITLES:
        # 精确匹配或以特殊标题开头且剩余≤1字符（排除"后记》："等）
        if text_compact == st or (text_compact.startswith(st) and len(text_compact) <= len(st) + 1):
            return True
    return False


def is_chapter_title(text, max_sz, bbox):
    """检测章标题（章扉页上的大字号标题）"""
    if max_sz < CHAPTER_MIN_SZ:
        return False
    if len(text) < 4 or len(text) > 25:
        return False
    w = bbox[2] - bbox[0]
    if w > 1000:
        return False
    # 排除含冒号的（那是节标题）
    if '：' in text:
        return False
    # 排除以数字开头的
    if text[0].isdigit():
        return False
    return True


def is_image_caption(item, prev_img_bbox):
    """检测图片说明文字：字号小 + 在图片下方附近"""
    if prev_img_bbox is None:
        return False
    if item['sz'] > IMG_CAPTION_MAX_SZ:
        return False
    # 文字在图片下方（y_start > img_y_end - 20）
    if item['bbox'][1] < prev_img_bbox[3] - 20:
        return False
    # 文字在图片下方不远（y_start < img_y_end + gap）
    if item['bbox'][1] > prev_img_bbox[3] + IMG_CAPTION_Y_GAP:
        return False
    return True


def clean_text(text):
    """清理文字：去除多余空格、OCR噪声"""
    text = text.strip()
    return text


def clean_noise(text):
    """去除OCR噪声：嵌入汉字间的短ASCII、孤立符号等"""
    if not text:
        return text
    # 去除汉字之间的短ASCII噪声（1-3字母），如"缺 Ml乏"→"缺乏"
    text = re.sub(r'(?<=[\u4e00-\u9fff])\s*[A-Za-z]{1,3}\s*(?=[\u4e00-\u9fff])', '', text)
    # 去除行首紧邻汉字的短ASCII噪声，如"Bl 正学小学简介"→"正学小学简介"
    text = re.sub(r'^[A-Za-z]{1,3}\s+(?=[\u4e00-\u9fff])', '', text)
    # 去除连续符号噪声（如"■ ■"、"□□"）
    text = re.sub(r'[■□▪▫▲△◆◇※]{2,}', '', text)
    return text.strip()


def normalize_title(text):
    """规范化标题文字：去除前缀噪声"""
    text = text.strip()
    # 去除前导单个字母（如 'a养病：...' → '养病：...'）
    match = re.match(r'^[a-zA-Z](.+)', text)
    if match and '：' in match.group(1):
        text = match.group(1)
    return text


# ============================================================
# 页面内容提取
# ============================================================
def _boxes_adjacent(a, b, gap=50.0):
    """判断两个bbox是否相邻或重叠"""
    return not (a[2] + gap < b[0] or b[2] + gap < a[0]
                or a[3] + gap < b[1] or b[3] + gap < a[1])


def merge_image_bboxes(bboxes, gap=50.0):
    """合并相邻的图片bbox（处理OCR将整图分割成碎片的问题）"""
    boxes = [tuple(bb) for bb in bboxes]
    changed = True
    while changed:
        changed = False
        result = []
        while boxes:
            bb = boxes.pop(0)
            i = 0
            while i < len(boxes):
                if _boxes_adjacent(bb, boxes[i], gap):
                    b2 = boxes.pop(i)
                    bb = (min(bb[0], b2[0]), min(bb[1], b2[1]),
                          max(bb[2], b2[2]), max(bb[3], b2[3]))
                    changed = True
                else:
                    i += 1
            result.append(bb)
        boxes = result
    return boxes


def extract_page_items(page, page_num):
    """提取页面内容项，按y坐标排序（图片碎片自动合并）"""
    d = page.get_text('dict')
    raw_text = []  # (bbox, text, max_sz)
    img_bboxes = []
    for b in d['blocks']:
        bb = b['bbox']
        if b['type'] == 1:  # 图片block
            if is_fullpage_image(bb):
                continue
            iw = bb[2] - bb[0]
            ih = bb[3] - bb[1]
            if iw < IMG_MIN_W or ih < IMG_MIN_H:
                continue
            if iw * ih < IMG_MIN_AREA:
                continue
            ratio = iw / ih if ih > 0 else 999
            if ratio > IMG_MAX_RATIO or ratio < 1.0 / IMG_MAX_RATIO:
                continue
            img_bboxes.append(bb)
        else:  # 文字block
            if is_footer(bb) or is_header(bb):
                continue
            spans = [s for l in b['lines'] for s in l['spans']]
            text = ''.join(s['text'] for s in spans).strip()
            if not text:
                continue
            if is_right_edge_deco(bb, text):
                continue
            max_sz = max(s['size'] for s in spans) if spans else 0
            raw_text.append((bb, text, max_sz))

    # 检测图片文字页：页面大部分文字字号远小于正文 → 整页截为图片，不提取文字
    if len(raw_text) >= 5:
        big = sum(1 for (_bb, _t, sz) in raw_text if sz >= IMG_TEXT_BIG_SZ)
        if big / len(raw_text) < IMG_TEXT_BIG_RATIO:
            all_bb = [bb for (bb, _t, _sz) in raw_text] + img_bboxes
            xs0 = [bb[0] for bb in all_bb]
            ys0 = [bb[1] for bb in all_bb]
            xs1 = [bb[2] for bb in all_bb]
            ys1 = [bb[3] for bb in all_bb]
            clip = (max(0.0, min(xs0) - 10), max(0.0, min(ys0) - 10),
                    min(PAGE_W, max(xs1) + 10), min(PAGE_H, max(ys1) + 10))
            return [{
                'type': 'img',
                'bbox': clip,
                'page': page_num,
                'imagetext': True,
            }]

    # 非图片文字页：应用噪声过滤
    items = []
    for (bb, text, max_sz) in raw_text:
        # 过滤小字号OCR噪声（如"HL""IW"等，正文46-50，图片说明35）
        if max_sz < NOISE_MIN_SZ:
            continue
        # 过滤纯短ASCII噪声块（如"Bl""Ml""OSS"等）
        t_strip = text.replace(' ', '').replace('\u3000', '')
        if len(t_strip) <= 3 and t_strip.isascii() and t_strip.isalpha():
            continue
        items.append({
            'type': 'text',
            'bbox': bb,
            'text': text,
            'sz': max_sz,
            'page': page_num,
        })
    # 合并相邻图片碎片
    for mb in merge_image_bboxes(img_bboxes):
        items.append({
            'type': 'img',
            'bbox': mb,
            'page': page_num,
        })
    # 按y坐标排序
    items.sort(key=lambda x: (x['bbox'][1], x['bbox'][0]))
    return items


def detect_title(item):
    """检测标题类型：section/special/None（章标题改用页脚提取）"""
    text = item['text']
    sz = item['sz']
    bb = item['bbox']
    if is_special_title(text, sz, bb):
        return 'special'
    if is_section_title(text, sz, bb):
        return 'section'
    return None


# ============================================================
# 页脚章标题提取
# ============================================================
def _clean_footer_label(footer):
    """清洗页脚标签：去除页码数字、噪声符号"""
    text = footer
    # 去除页码模式（如 "0 0 3", "0 1 0", "2 0" 等）
    text = re.sub(r'\d\s*\d\s*\d', '', text)
    text = re.sub(r'\d\s*\d(?=\s|$)', '', text)
    # 去除末尾/开头的数字
    text = re.sub(r'\d+\s*$', '', text)
    text = re.sub(r'^\s*\d+', '', text)
    # 去除孤立符号和ASCII噪声
    text = re.sub(r'[■£E|S\]\[U!o]', '', text)
    # 去除汉字间的单个数字（如"包3"→"包"）
    text = re.sub(r'(?<=[\u4e00-\u9fff])\d+(?=[\u4e00-\u9fff\s])', '', text)
    # 去除汉字间的单个ASCII字母噪声
    text = re.sub(r'(?<=[\u4e00-\u9fff])[A-Za-z](?=[\u4e00-\u9fff])', '', text)
    # 去除末尾的中文数字（OCR噪声，如"梦四"→"梦"）
    text = re.sub(r'[一二三四五六七八九十]+$', '', text)
    return text.strip()


def _chapter_key(title):
    """提取章标题的关键部分用于匹配：去除空格、中文数字前缀"""
    t = title.replace(' ', '').replace('\u3000', '')
    t = re.sub(r'^[一二三四五六七八九十]+', '', t)
    return t


def extract_footer_chapters(doc):
    """从页脚提取章标题和书名

    奇数页页脚格式: {章标题}{页码}
    偶数页页脚格式: {页码}{书名}

    返回:
        page_chapters: {page_idx(0-based): chapter_key}
        chapter_order: [chapter_key] 按首次出现顺序
        book_name: str
    """
    raw_labels = {}  # page_idx → cleaned footer text
    book_name = None

    for i in range(doc.page_count):
        page = doc[i]
        d = page.get_text('dict')
        footer_texts = []
        for b in d['blocks']:
            if b['type'] == 0:
                bb = b['bbox']
                if bb[1] > FOOT_Y_THRESHOLD:
                    spans = [s for l in b['lines'] for s in l['spans']]
                    text = ''.join(s['text'] for s in spans).strip()
                    if text:
                        footer_texts.append(text)
        if not footer_texts:
            continue
        footer = ' '.join(footer_texts)

        # 如果包含"柔石画传"→ 书名（偶数页）
        if '柔石画传' in footer:
            book_name = '柔石画传'
            continue

        # 清洗
        cleaned = _clean_footer_label(footer)
        if not cleaned or len(cleaned) < 4:
            continue

        raw_labels[i] = cleaned

    # 模糊聚类归一化章标题
    chapter_groups = []  # [(key, [page_indices])]
    for page_idx in sorted(raw_labels.keys()):
        label = raw_labels[page_idx]
        key = _chapter_key(label)

        matched = False
        for grp in chapter_groups:
            ratio = SequenceMatcher(None, key, grp[0]).ratio()
            if ratio > 0.7:
                grp[1].append(page_idx)
                # 取更长的key作为代表
                if len(key) > len(grp[0]):
                    grp[0] = key
                matched = True
                break
        if not matched:
            chapter_groups.append([key, [page_idx]])

    # 过滤噪声章标题
    def _is_chapter_noise(key):
        """检测是否是噪声章标题"""
        # 中文字符占比<50%
        if _chinese_ratio(key) < 0.5:
            return True
        # 匹配 SPECIAL_TITLES（这些在 build_catalog 中单独处理）
        for st in SPECIAL_TITLES:
            st_compact = st.replace(' ', '')
            if st_compact in key or key in st_compact:
                return True
        # 含URL/价格关键词
        if any(kw in key for kw in ['www', 'http', '定价', 'ISBN', '.com', '.cn', 'ewen']):
            return True
        return False

    chapter_groups = [g for g in chapter_groups if not _is_chapter_noise(g[0]) and len(g[1]) >= 3]

    # 按首次出现的页码排序
    chapter_groups.sort(key=lambda g: g[1][0])

    # 建立页→章key映射
    page_chapters = {}
    chapter_order = []
    for key, page_indices in chapter_groups:
        chapter_order.append(key)
        for pi in page_indices:
            page_chapters[pi] = key

    return page_chapters, chapter_order, book_name


# ============================================================
# 图片提取
# ============================================================
def extract_image(page, bbox, out_path):
    """从页面截取指定区域为PNG图片"""
    clip = pymupdf.Rect(bbox)
    # 稍微扩大裁剪区域以避免边缘截断
    clip = pymupdf.Rect(
        max(0, clip.x0 - 2),
        max(0, clip.y0 - 2),
        min(PAGE_W, clip.x1 + 2),
        min(PAGE_H, clip.y1 + 2),
    )
    pix = page.get_pixmap(clip=clip, alpha=False)
    pix.save(str(out_path))
    return out_path


# ============================================================
# 节处理
# ============================================================
class Section:
    """一个节（对应一个txt文件）"""
    def __init__(self, title, level=3):
        self.title = title          # 节标题
        self.level = level          # 层级（2=章, 3=节）
        self.lines = []             # txt行列表
        self.img_count = 0          # 图片计数
        self.pages = []             # 涉及的页码

    def add_text(self, text):
        """添加段落文字"""
        text = clean_noise(text)
        if text.strip():
            self.lines.append(text.strip())

    def add_image(self, img_path):
        """添加图片占位符"""
        self.lines.append(f'<img>{img_path}</img>')
        self.img_count += 1

    def add_caption(self, text):
        """添加图片说明（紧跟img，不空行）"""
        text = clean_noise(text)
        if text.strip():
            self.lines.append(text.strip())

    def add_blank(self):
        """添加空行"""
        self.lines.append('')


# ============================================================
# 主处理逻辑
# ============================================================
def process_pdf(pdf_path, output_dir):
    """主处理函数"""
    pdf_path = Path(pdf_path)
    output_dir = Path(output_dir)
    txt_dir = output_dir / 'txt'
    img_dir = output_dir / 'img'
    txt_dir.mkdir(parents=True, exist_ok=True)
    img_dir.mkdir(parents=True, exist_ok=True)

    doc = pymupdf.open(str(pdf_path))

    # === 从页脚提取章标题和书名 ===
    page_chapters, chapter_order, footer_book_name = extract_footer_chapters(doc)
    book_name = footer_book_name if footer_book_name else pdf_path.stem
    print(f'书名: {book_name}')
    print(f'检测到 {len(chapter_order)} 个章标题:')
    for idx, ck in enumerate(chapter_order):
        cn = CN_NUMS[idx] if idx < len(CN_NUMS) else str(idx + 1)
        print(f'  {cn} {ck}')

    # === 第一遍：检测所有标题，建立节结构 ===
    all_items = []  # (page_num, items)
    title_marks = []  # (page_num, item_index, title_type, title_text)

    # 目录页范围（检测到"目录"标题后，跳过后续目录页的节标题检测）
    toc_pages = set()

    for i in range(doc.page_count):
        page = doc[i]
        items = extract_page_items(page, i + 1)
        all_items.append(items)
        for j, item in enumerate(items):
            if item['type'] == 'text':
                ttype = detect_title(item)
                if ttype:
                    title_text = normalize_title(item['text'])
                    title_marks.append((i, j, ttype, title_text))
                    # 检测到"目录"标题 → 标记后续4页为目录页
                    if ttype == 'special' and '目录' in title_text.replace(' ', ''):
                        for k in range(i, min(i + 5, doc.page_count)):
                            toc_pages.add(k)

    # 过滤掉目录页内的所有标题（保留"目录"标题本身作为节）
    title_marks = [
        (pg, it_idx, ttype, ttext)
        for (pg, it_idx, ttype, ttext) in title_marks
        if not (pg in toc_pages and '目录' not in ttext.replace(' ', ''))
    ]

    # === 建立节列表 ===
    sections_info = []

    for idx, (pg, it_idx, ttype, ttext) in enumerate(title_marks):
        if ttype in ('section', 'special'):
            sections_info.append({
                'page': pg,
                'item_idx': it_idx,
                'title': ttext,
                'level': 3,
            })

    # 如果没有检测到任何节标题，将整个文档作为一个节
    if not sections_info:
        sections_info.append({
            'page': 0,
            'item_idx': -1,
            'title': book_name,
            'level': 3,
        })

    # === 第二遍：按节提取内容 ===
    sections = []  # [(sec_num, Section, page_idx)]
    current_section = None
    section_counter = 0

    # 建立页→节起始映射
    section_starts = {}  # page → [section_info]
    for si in sections_info:
        section_starts.setdefault(si['page'], []).append(si)

    for i in range(doc.page_count):
        page = doc[i]
        items = all_items[i]

        # 目录页：只处理节标题起始，跳过其他所有内容
        is_toc_page = i in toc_pages

        # 检查是否有节标题
        new_sections_on_page = []
        if i in section_starts:
            for si in section_starts[i]:
                new_sections_on_page.append(si)

        # 处理页面items
        para_buffer = []  # 当前段落缓冲（同页连续文字block合并）
        prev_img_bbox = None
        section_started_this_page = False

        for j, item in enumerate(items):
            # 目录页：跳过所有非节标题起始的内容
            if is_toc_page:
                is_sec_start = any(j == si['item_idx'] for si in new_sections_on_page)
                if not is_sec_start:
                    continue

            # 检查是否是节标题
            is_new_section = False
            for si in new_sections_on_page:
                if j == si['item_idx']:
                    # 先flush段落缓冲
                    if para_buffer and current_section:
                        current_section.add_text(' '.join(para_buffer))
                        para_buffer = []
                    # 创建新节
                    section_counter += 1
                    sec_num = f'{section_counter:04d}'
                    current_section = Section(si['title'], level=3)
                    current_section.pages.append(i + 1)
                    sections.append((sec_num, current_section, i))
                    section_started_this_page = True
                    is_new_section = True
                    break

            if is_new_section:
                continue

            if item['type'] == 'img':
                # 先flush段落缓冲
                if para_buffer and current_section:
                    current_section.add_text(' '.join(para_buffer))
                    current_section.add_blank()
                    para_buffer = []
                # 提取图片
                if current_section is None:
                    # 还没有节 → 前置封面内容，跳过
                    continue

                sec_num = sections[-1][0]
                img_name = f'{sec_num}_{current_section.img_count + 1}.png'
                img_path = img_dir / img_name
                try:
                    extract_image(page, item['bbox'], img_path)
                    current_section.add_image(f'img/{img_name}')
                except Exception as e:
                    print(f'  [警告] 图片提取失败 第{i+1}页: {e}')
                prev_img_bbox = item['bbox']

            elif item['type'] == 'text':

                # 检查是否是图片说明
                if is_image_caption(item, prev_img_bbox):
                    if current_section:
                        current_section.add_caption(item['text'])
                        current_section.add_blank()
                    prev_img_bbox = None
                    continue

                # 检查是否是特殊标题（已作为节标题处理）
                if ttype == 'special':
                    # 如果不是节起始，跳过
                    is_start = False
                    for si in new_sections_on_page:
                        if j == si['item_idx']:
                            is_start = True
                            break
                    if is_start:
                        continue
                    # 否则可能是重复文字，跳过
                    continue

                # 检查是否是节标题（已作为节起始处理）
                if ttype == 'section':
                    continue

                # 普通文字
                if current_section is None:
                    # 还没有节 → 前置封面内容，跳过
                    continue

                para_buffer.append(item['text'])
                prev_img_bbox = None

        # 页面结束，flush段落缓冲
        if para_buffer and current_section:
            current_section.add_text(' '.join(para_buffer))
            current_section.add_blank()

    # === 输出txt文件 ===
    print(f'\n=== 输出txt文件 ===')
    for sec_num, section, page_idx in sections:
        txt_path = txt_dir / f'{sec_num}.txt'
        # 构建txt内容：首行标题，空行，然后内容
        lines = [section.title, '']
        lines.extend(section.lines)
        # 去除尾部多余空行
        while lines and lines[-1] == '':
            lines.pop()
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n')
        print(f'  {sec_num}.txt: {section.title[:30]} ({section.img_count}图, {len(section.lines)}行)')

    # === 生成目录.xlsx ===
    print(f'\n=== 生成目录.xlsx ===')
    build_catalog(book_name, sections, page_chapters, chapter_order, output_dir / '目录.xlsx')

    doc.close()
    print(f'\n完成！共 {len(sections)} 节, {sum(s[1].img_count for s in sections)} 张图片')
    return sections


def build_catalog(book_name, sections, page_chapters, chapter_order, out_path):
    """生成目录.xlsx

    结构：书名(1) → 书名(2) → 前置特殊标题(3) → 目录(2)→(3) → 章(2)→节(3) → 后置特殊标题(2)→(3)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'sheet1'
    ws.append(['序号', '目录名', '层级', '整理后TXT'])

    row_num = 1
    # 层级1：书名
    ws.append([row_num, book_name, 1, None])
    row_num += 1
    # 层级2：书名（重复）
    ws.append([row_num, book_name, 2, None])
    row_num += 1

    # 分类节：前置、目录、正文（章→节）、后置
    PRE_SPECIAL = ['龙华英烈画传系列丛书编委会', '出版说明']
    POST_SPECIAL = ['柔石大事年表', '参考文献', '后记']

    pre_sections = []
    toc_section = None
    body_sections = []
    post_sections = []

    for sec_num, section, page_idx in sections:
        title_compact = section.title.replace(' ', '').replace('\u3000', '')
        is_pre = any(title_compact.startswith(ps.replace(' ', '')) for ps in PRE_SPECIAL)
        is_post = any(title_compact.startswith(ps.replace(' ', '')) for ps in POST_SPECIAL)
        is_toc = '目录' in title_compact

        if is_pre:
            pre_sections.append((sec_num, section))
        elif is_toc:
            toc_section = (sec_num, section)
        elif is_post:
            post_sections.append((sec_num, section))
        else:
            body_sections.append((sec_num, section, page_idx))

    # 输出前置特殊标题（层级3，在书名层级2下）
    for sec_num, section in pre_sections:
        ws.append([row_num, section.title, 3, sec_num])
        row_num += 1

    # 输出目录（层级2+3）
    if toc_section:
        sec_num, section = toc_section
        ws.append([row_num, section.title, 2, None])
        row_num += 1
        ws.append([row_num, section.title, 3, sec_num])
        row_num += 1

    # 为正文节分配章：找到每个章的首次出现页，构建页码范围
    chapter_first_page = {}
    for pg in sorted(page_chapters.keys()):
        key = page_chapters[pg]
        if key not in chapter_first_page:
            chapter_first_page[key] = pg

    chapter_ranges = []
    for idx, ch_key in enumerate(chapter_order):
        start = chapter_first_page[ch_key]
        if idx + 1 < len(chapter_order):
            end = chapter_first_page[chapter_order[idx + 1]]
        else:
            end = 9999
        chapter_ranges.append((start, end, ch_key))

    # 为每个节找到所属章
    sections_by_chapter = {}
    for sec_num, section, page_idx in body_sections:
        ch_key = None
        for start, end, key in chapter_ranges:
            if start <= page_idx < end:
                ch_key = key
                break
        if ch_key:
            sections_by_chapter.setdefault(ch_key, []).append((sec_num, section))

    # 输出章→节（按chapter_order顺序）
    for ch_idx, ch_key in enumerate(chapter_order):
        if ch_key not in sections_by_chapter:
            continue
        cn = CN_NUMS[ch_idx] if ch_idx < len(CN_NUMS) else str(ch_idx + 1)
        ws.append([row_num, f'{cn} {ch_key}', 2, None])
        row_num += 1
        for sec_num, section in sections_by_chapter[ch_key]:
            ws.append([row_num, section.title, 3, sec_num])
            row_num += 1

    # 输出后置特殊标题（层级2+3）
    for sec_num, section in post_sections:
        ws.append([row_num, section.title, 2, None])
        row_num += 1
        ws.append([row_num, section.title, 3, sec_num])
        row_num += 1

    # 设置列宽
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 6
    ws.column_dimensions['D'].width = 12

    wb.save(str(out_path))
    print(f'  目录.xlsx: {row_num - 1} 行')


# ============================================================
# 入口
# ============================================================
def main():
    if len(sys.argv) < 2:
        print('用法: python extract_book.py <pdf路径> [输出目录]')
        print('示例: python extract_book.py NH0633.pdf 柔石画传')
        sys.exit(1)

    pdf_path = sys.argv[1]
    if len(sys.argv) >= 3:
        output_dir = sys.argv[2]
    else:
        output_dir = Path(pdf_path).stem

    print(f'PDF: {pdf_path}')
    print(f'输出目录: {output_dir}')

    process_pdf(pdf_path, output_dir)


if __name__ == '__main__':
    main()