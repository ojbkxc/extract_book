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
import re
import pymupdf
import openpyxl
from pathlib import Path
from difflib import SequenceMatcher

# ============================================================
# 配置参数（默认值，实际运行时由 PdfConfig 自适应覆盖）
# ============================================================
# 通用特殊标题（大多数图书都有的前置/后置部分）
COMMON_SPECIAL_TITLES = [
    '出版说明',
    '目录',
    '前言',
    '序言',
    '序',
    '后记',
    '附录',
    '参考文献',
    '大事年表',
]

# 中文数字（章标题前缀）
CN_NUMS = '一二三四五六七八九十'

_CN_DIGITS = '零一二三四五六七八九'


def to_cn_num(n):
    """将 1-99 的整数转换为中文数字

    示例: 10→十, 11→十一, 20→二十, 99→九十九。
    超出 1-99 支持范围时回退为阿拉伯数字字符串。
    """
    if n < 1:
        return str(n)
    if n < 10:
        return _CN_DIGITS[n]
    if n < 20:
        return '十' + (_CN_DIGITS[n - 10] if n > 10 else '')
    tens, ones = divmod(n, 10)
    if tens >= 10:  # 超出 1-99 支持范围
        return str(n)
    if ones == 0:
        return _CN_DIGITS[tens] + '十'
    return _CN_DIGITS[tens] + '十' + _CN_DIGITS[ones]


class PdfConfig:
    """PDF自适应配置：从实际PDF推断页面尺寸、字号分布等参数"""

    def __init__(self, doc, special_titles=None, page_dicts=None):
        self.special_titles = list(COMMON_SPECIAL_TITLES)
        if special_titles:
            self.special_titles.extend(special_titles)
        self._analyze(doc, page_dicts=page_dicts)

    def _analyze(self, doc, page_dicts=None):
        # === 页面尺寸 ===
        page0 = doc[0]
        self.page_w = page0.rect.width
        self.page_h = page0.rect.height

        # === 字号分布统计（采样前50页）===
        sz_counter = {}
        for i in range(min(doc.page_count, 50)):
            d = page_dicts[i] if page_dicts is not None else doc[i].get_text('dict')
            for b in d['blocks']:
                if b['type'] == 0:
                    for l in b['lines']:
                        for s in l['spans']:
                            sz = round(s['size'])
                            sz_counter[sz] = sz_counter.get(sz, 0) + len(s['text'].strip())

        # 最频繁的字号 = 正文字号
        self.body_sz = max(sz_counter, key=sz_counter.get) if sz_counter else 50.0

        # === 页眉页脚阈值（按页面高度比例）===
        self.foot_y = self.page_h * 0.90
        self.head_y = self.page_h * 0.07
        self.right_edge_x = self.page_w * 0.88
        self.right_edge_max_w = 100.0

        # === 全页背景图 ===
        self.fullpage_w_ratio = 0.95
        self.fullpage_h_ratio = 0.95

        # === 节标题检测 ===
        self.section_min_sz = self.body_sz + 2  # 正文字号+2（节标题字号略高于正文）
        self.section_max_width = self.page_w * 0.52
        self.section_min_len = 4
        self.section_max_len = 30


        # === 图片说明 ===
        self.img_caption_max_sz = self.body_sz * 0.78
        self.img_caption_y_gap = self.body_sz * 1.2

        # === 小图片过滤 ===
        self.img_min_w = 50.0
        self.img_min_h = 50.0
        self.img_min_area = (self.page_w * self.page_h) * 0.01  # 页面面积的1%
        self.img_max_ratio = 8.0

        # === OCR噪声过滤 ===
        self.noise_min_sz = self.body_sz * 0.55

        # === 图片文字页检测 ===
        self.img_text_big_sz = self.body_sz * 0.82
        self.img_text_big_ratio = 0.2

        print(f'PDF自适应: 页面{self.page_w:.0f}×{self.page_h:.0f}, 正文字号{self.body_sz:.0f}')
        print(f'  页脚阈值={self.foot_y:.0f}, 节标题字号≥{self.section_min_sz:.0f}, 噪声字号<{self.noise_min_sz:.0f}')


# 全局配置（在 process_pdf 中初始化）
_cfg = None


# ============================================================
# 过滤检测函数
# ============================================================
def is_fullpage_image(bbox):
    """检测全页背景图"""
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]
    return w > _cfg.page_w * _cfg.fullpage_w_ratio and h > _cfg.page_h * _cfg.fullpage_h_ratio


def is_footer(bbox):
    """检测页脚"""
    return bbox[1] > _cfg.foot_y


def is_header(bbox):
    """检测页眉"""
    return bbox[3] < _cfg.head_y


def is_right_edge_deco(bbox, text=''):
    """检测右侧竖排装饰文字（保留含'目录'的特殊标题）"""
    w = bbox[2] - bbox[0]
    if bbox[0] > _cfg.right_edge_x and w < _cfg.right_edge_max_w:
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
    """检测节标题：含'：' + 字号大 + 宽度合理 + 长度4-30 + 非数字开头; 中文字符占比>60%"""
    w = bbox[2] - bbox[0]
    if '：' not in text:
        return False
    if max_sz < _cfg.section_min_sz:
        return False
    if w > _cfg.section_max_width:
        return False
    if len(text) < _cfg.section_min_len or len(text) > _cfg.section_max_len:
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
    # 排除以引号结尾且引号前是句号/感叹号的（如"...执行！""）
    if text[-1] in '”"』』' and len(text) >= 2 and text[-2] in '。！':
        return False
    # 冒号前后部分
    parts = text.split('：', 1)
    prefix = parts[0]
    suffix = parts[1] if len(parts) > 1 else ''
    if len(prefix) < 1 or len(prefix) > 8:
        return False
    # 排除前缀含逗号的（如"的信中，他写道："是正文，节标题前缀不含逗号）
    if '，' in prefix or ',' in prefix:
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
    if max_sz < _cfg.body_sz * 0.9:
        return False
    w = bbox[2] - bbox[0]
    if w > _cfg.page_w * 0.42:
        return False
    # 去除空格后模糊匹配
    text_compact = text.replace(' ', '').replace('\u3000', '')
    for st in _cfg.special_titles:
        # 精确匹配或以特殊标题开头且剩余≤1字符（排除"后记》："等）
        if text_compact == st or (text_compact.startswith(st) and len(text_compact) <= len(st) + 1):
            return True
    return False



def is_image_caption(item, prev_img_bbox):
    """检测图片说明文字：字号小 + 在图片下方附近"""
    if prev_img_bbox is None:
        return False
    if item['sz'] > _cfg.img_caption_max_sz:
        return False
    # 文字在图片下方（y_start > img_y_end - 20）
    if item['bbox'][1] < prev_img_bbox[3] - 20:
        return False
    # 文字在图片下方不远（y_start < img_y_end + gap）
    if item['bbox'][1] > prev_img_bbox[3] + _cfg.img_caption_y_gap:
        return False
    return True



def clean_noise(text):
    """去除OCR噪声：嵌入汉字间的短ASCII、孤立符号等

    汉字之间的 1-3 个字母仅当为"单个字母"或"大小写混杂"（OCR噪声特征，
    如 Ml/Bl/HL）时删除；全大写或全小写的字母串（如 WTO/GPS/DNA 术语）保留。
    """
    if not text:
        return text

    def _drop_ocr_letter(m):
        latin = m.group('latin')
        # 全大写/全小写 → 视为术语，原样保留；单个字母或大小写混杂 → OCR噪声，删除
        if len(latin) > 1 and (latin.isupper() or latin.islower()):
            return m.group(0)
        return ''

    # 去除汉字之间的短ASCII噪声（1-3字母），如"缺 Ml乏"→"缺乏"，但保留"加入WTO后的影响"
    text = re.sub(
        r'(?<=[\u4e00-\u9fff])\s*(?P<latin>[A-Za-z]{1,3})\s*(?=[\u4e00-\u9fff])',
        _drop_ocr_letter, text)
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


def extract_page_items(page, page_num, d=None):
    """提取页面内容项，按y坐标排序（图片碎片自动合并）

    d: 可选的已获取的 get_text('dict') 结果，避免同页重复解析。
    """
    if d is None:
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
            if iw < _cfg.img_min_w or ih < _cfg.img_min_h:
                continue
            if iw * ih < _cfg.img_min_area:
                continue
            ratio = iw / ih if ih > 0 else 999
            if ratio > _cfg.img_max_ratio or ratio < 1.0 / _cfg.img_max_ratio:
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
        big = sum(1 for (_bb, _t, sz) in raw_text if sz >= _cfg.img_text_big_sz)
        if big / len(raw_text) < _cfg.img_text_big_ratio:
            all_bb = [bb for (bb, _t, _sz) in raw_text] + img_bboxes
            xs0 = [bb[0] for bb in all_bb]
            ys0 = [bb[1] for bb in all_bb]
            xs1 = [bb[2] for bb in all_bb]
            ys1 = [bb[3] for bb in all_bb]
            clip = (max(0.0, min(xs0) - 10), max(0.0, min(ys0) - 10),
                    min(_cfg.page_w, max(xs1) + 10), min(_cfg.page_h, max(ys1) + 10))
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
        if max_sz < _cfg.noise_min_sz:
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


def extract_footer_chapters(doc, page_dicts=None):
    """从页脚提取章标题和书名

    奇数页页脚格式: {章标题}{页码}
    偶数页页脚格式: {页码}{书名}

    书名通过频率统计自动识别（在大量页中反复出现的相同文字）

    返回:
        page_chapters: {page_idx(0-based): chapter_key}
        chapter_order: [chapter_key] 按首次出现顺序
        book_name: str
    """
    from collections import Counter

    # 先收集所有页脚标签（清洗后）
    all_footer_labels = {}  # page_idx → cleaned label
    for i in range(doc.page_count):
        d = page_dicts[i] if page_dicts is not None else doc[i].get_text('dict')
        footer_texts = []
        for b in d['blocks']:
            if b['type'] == 0:
                bb = b['bbox']
                if bb[1] > _cfg.foot_y:
                    spans = [s for l in b['lines'] for s in l['spans']]
                    text = ''.join(s['text'] for s in spans).strip()
                    if text:
                        footer_texts.append(text)
        if not footer_texts:
            continue
        footer = ' '.join(footer_texts)
        cleaned = _clean_footer_label(footer)
        if cleaned and len(cleaned) >= 2:
            all_footer_labels[i] = cleaned

    # 统计标签出现频率，自动识别书名（出现频率最高的标签）
    label_counter = Counter(all_footer_labels.values())
    total_pages = doc.page_count

    book_name = None
    book_name_label = None
    for label, count in label_counter.most_common():
        # 书名特征：在 >= 20% 的页中出现，且长度 >= 2
        if count >= total_pages * 0.2 and len(label) >= 2:
            book_name = label
            book_name_label = label
            break

    # 章标题候选 = 排除书名页后的标签
    raw_labels = {}
    for page_idx, label in all_footer_labels.items():
        if book_name_label and label == book_name_label:
            continue
        if len(label) < 4:
            continue
        raw_labels[page_idx] = label

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
        # 匹配 _cfg.special_titles（这些在 build_catalog 中单独处理）
        for st in _cfg.special_titles:
            st_compact = st.replace(' ', '')
            if st_compact in key or key in st_compact:
                return True
        # 含URL/域名/价格等非正文关键词（通用判断，不依赖具体书名）
        low = key.lower()
        if any(kw in low for kw in ['www', 'http', '.com', '.cn', 'isbn']):
            return True
        if '定价' in key:
            return True
        # 含'.'（域名特征）
        if '.' in key:
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
# 目录页检测与章标题提取
# ============================================================
def _block_text(block):
    """拼接一个文字 block 中所有 span 的文字"""
    return ''.join(s['text'] for l in block['lines'] for s in l['spans'])


def _page_is_toc(d, threshold=0.5):
    """判断单页是否符合目录页特征：大部分文字 block 以页码数字结尾"""
    total = 0
    end_digit = 0
    for b in d['blocks']:
        if b['type'] != 0:
            continue
        bb = b['bbox']
        if is_footer(bb) or is_header(bb):
            continue
        text = _block_text(b).strip()
        if not text:
            continue
        total += 1
        if re.search(r'\d\s*$', text):
            end_digit += 1
    return total > 0 and (end_digit / total) >= threshold


def _is_toc_following_divider(d):
    """判断目录列表之后紧随的章扉页/分隔页

    特征：文字 block 极少（1-4）且总字符数很小，通常为竖排章名或纯图装饰。
    这类页在手工整理时归入目录区、不参与正文提取，需并入目录页范围。
    """
    text_blocks = 0
    text_chars = 0
    for b in d['blocks']:
        if b['type'] != 0:
            continue
        bb = b['bbox']
        if is_footer(bb) or is_header(bb):
            continue
        t = _block_text(b).strip()
        if t:
            text_blocks += 1
            text_chars += len(t)
    return 1 <= text_blocks <= 4 and text_chars <= 40


def detect_toc_pages(doc, page_dicts=None):
    """动态检测目录页范围（返回 0-based 页码集合）

    规则：从含"目录"标题的页开始，连续满足"大部分文字 block 以页码数字结尾"
    的页视为目录页；遇到首个不满足的页即停止。若停止处恰为紧随目录的章扉页/
    分隔页（文字极少），一并纳入目录区后再停止。
    """
    toc_pages = set()
    started = False
    i = 0
    n = doc.page_count
    while i < n:
        d = page_dicts[i] if page_dicts is not None else doc[i].get_text('dict')
        if not started:
            for b in d['blocks']:
                if b['type'] != 0:
                    continue
                compact = _block_text(b).replace(' ', '').replace('\u3000', '').strip()
                if compact == '目录':
                    started = True
                    toc_pages.add(i)
                    break
            i += 1
            continue
        if _page_is_toc(d):
            toc_pages.add(i)
            i += 1
            continue
        # 目录列表结束：紧随的章扉页/分隔页并入目录区，然后停止
        if _is_toc_following_divider(d):
            toc_pages.add(i)
        break
    return toc_pages


# 目录行清洗正则
_TOC_PAGE_NUM_RE = re.compile(r'[\s。.、]+\d[\d\s。.、]*$')
_TOC_SYMBOL_NOISE_RE = re.compile(r'\s*[\^■□▪▫▲△◆◇※]+\s*\d*')
_TOC_SHORT_ASCII_RE = re.compile(r'(?:(?<=[\u4e00-\u9fff])|\s)[A-Za-z]{1,4}(?![A-Za-z])')


def _clean_toc_line(text):
    """清洗目录行：去除末尾页码与 OCR 噪声"""
    t = text.strip()
    t = _TOC_PAGE_NUM_RE.sub('', t)       # 去除末尾页码（含空格分隔，如 " 1 15"）
    t = re.sub(r'\d{1,4}\s*$', '', t)     # 兜底：去除紧贴的末尾数字
    t = _TOC_SYMBOL_NOISE_RE.sub('', t)   # 去除 ^■9 等符号噪声
    t = _TOC_SHORT_ASCII_RE.sub('', t)    # 去除 " It" 等短 ASCII 噪声
    return t.strip()


def extract_toc_chapters(doc, toc_pages, page_dicts=None):
    """从目录页提取有序的章标题文字列表

    章标题判定：去除末尾页码后，不含中文冒号"："、不属于 special_titles、
    中文字符占比 >= 0.6、长度 4-30；再用 _chapter_key 去掉中文数字前缀。
    """
    chapters = []
    seen = set()
    for i in sorted(toc_pages):
        d = page_dicts[i] if page_dicts is not None else doc[i].get_text('dict')
        for b in d['blocks']:
            if b['type'] != 0:
                continue
            bb = b['bbox']
            if is_footer(bb) or is_header(bb):
                continue
            text = _clean_toc_line(_block_text(b))
            if not text:
                continue
            # 含中文冒号 → 节标题，排除
            if '：' in text:
                continue
            # 属于特殊标题，排除
            compact = text.replace(' ', '').replace('\u3000', '')
            is_special = False
            for st in _cfg.special_titles:
                stc = st.replace(' ', '')
                if compact == stc or (compact.startswith(stc) and len(compact) <= len(stc) + 1):
                    is_special = True
                    break
            if is_special:
                continue
            if _chinese_ratio(text) < 0.6:
                continue
            if len(text) < 4 or len(text) > 30:
                continue
            key = _chapter_key(text)
            if not key or key in seen:
                continue
            seen.add(key)
            chapters.append(key)
    return chapters


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
        min(_cfg.page_w, clip.x1 + 2),
        min(_cfg.page_h, clip.y1 + 2),
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
def process_pdf(pdf_path, output_dir, special_titles=None):
    """主处理函数

    Args:
        pdf_path: PDF文件路径
        output_dir: 输出目录
        special_titles: 额外的特殊标题列表（如丛书编委会、大事年表等）
    """
    pdf_path = Path(pdf_path)
    output_dir = Path(output_dir)
    txt_dir = output_dir / 'txt'
    img_dir = output_dir / 'img'
    txt_dir.mkdir(parents=True, exist_ok=True)
    img_dir.mkdir(parents=True, exist_ok=True)

    # 清空输出目录的旧文件，避免重复运行时残留（批量场景尤为重要）
    for _d in (txt_dir, img_dir):
        for _f in _d.iterdir():
            if _f.is_file():
                _f.unlink()
    _old_xlsx = output_dir / '目录.xlsx'
    if _old_xlsx.exists():
        _old_xlsx.unlink()

    doc = pymupdf.open(str(pdf_path))

    # === 预取每页文字结构，供多处复用（性能优化）===
    page_dicts = [doc[i].get_text('dict') for i in range(doc.page_count)]

    # === 初始化自适应配置 ===
    global _cfg
    _cfg = PdfConfig(doc, special_titles=special_titles, page_dicts=page_dicts)

    # === 扫描件/无文字层检测 ===
    total_chars = sum(
        len(s['text'])
        for d in page_dicts
        for b in d['blocks'] if b['type'] == 0
        for l in b['lines'] for s in l['spans']
    )
    if total_chars < 100:
        print('  [警告] 全文文字总字符数过少(<100)，可能为扫描件或无文字层，提取结果可能为空')

    # === 动态检测目录页范围 ===
    toc_pages = detect_toc_pages(doc, page_dicts)
    if toc_pages:
        print(f'检测到目录页: 第 {min(toc_pages) + 1}-{max(toc_pages) + 1} 页')

    # === 从页脚提取章标题（用于节归属）和书名 ===
    page_chapters, chapter_order, footer_book_name = extract_footer_chapters(doc, page_dicts)
    book_name = footer_book_name if footer_book_name else pdf_path.stem
    print(f'书名: {book_name}')

    # === 从目录页提取章标题（主方案，显示文字）===
    toc_chapters = extract_toc_chapters(doc, toc_pages, page_dicts)
    if len(toc_chapters) >= 2 and len(toc_chapters) == len(chapter_order):
        display_chapters = toc_chapters
    else:
        display_chapters = chapter_order
        if len(toc_chapters) >= 2:
            print(f'  [提示] 目录章标题数({len(toc_chapters)})与页脚章数({len(chapter_order)})不一致，回退用页脚章标题')

    print(f'检测到 {len(display_chapters)} 个章标题:')
    for idx, ck in enumerate(display_chapters):
        print(f'  {to_cn_num(idx + 1)} {ck}')

    # === 第一遍：检测所有标题，建立节结构 ===
    all_items = []  # (page_num, items)
    title_marks = []  # (page_num, item_index, title_type, title_text)

    for i in range(doc.page_count):
        page = doc[i]
        items = extract_page_items(page, i + 1, page_dicts[i])
        all_items.append(items)
        for j, item in enumerate(items):
            if item['type'] == 'text':
                ttype = detect_title(item)
                if ttype:
                    title_text = normalize_title(item['text'])
                    title_marks.append((i, j, ttype, title_text))

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

                # 重新检测本 item 的标题类型（不要使用第一遍循环残留的 ttype 变量）
                ttype2 = detect_title(item)

                # 检查是否是特殊标题（已作为节标题处理）
                if ttype2 == 'special':
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
                if ttype2 == 'section':
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
    build_catalog(book_name, sections, page_chapters, chapter_order, display_chapters, output_dir / '目录.xlsx')

    doc.close()
    print(f'\n完成！共 {len(sections)} 节, {sum(s[1].img_count for s in sections)} 张图片')
    return sections


def build_catalog(book_name, sections, page_chapters, chapter_order, chapter_titles, out_path):
    """生成目录.xlsx

    结构：书名(1) → 书名(2) → 前置特殊标题(3) → 目录(2)→(3) → 章(2)→节(3) → 后置特殊标题(2)→(3)

    chapter_titles: 章标题显示文字列表（与 chapter_order 顺序一一对应）。
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
    # 有章：按页码位置分类；无章：按特殊标题名称分类
    has_chapters = bool(page_chapters)
    if has_chapters:
        first_chapter_page = min(page_chapters.keys())
        last_chapter_page = max(page_chapters.keys())
    else:
        first_chapter_page = 0
        last_chapter_page = 0

    # 无章书中判定为"后置"的特殊标题关键词
    post_keywords = ('后记', '附录', '参考文献', '大事年表')

    def _is_post_title(title_compact):
        return any(kw in title_compact for kw in post_keywords)

    pre_sections = []
    toc_section = None
    body_sections = []
    post_sections = []

    for sec_num, section, page_idx in sections:
        title_compact = section.title.replace(' ', '').replace('\u3000', '')
        is_toc = '目录' in title_compact
        is_special = any(
            st.replace(' ', '') in title_compact or title_compact.startswith(st.replace(' ', ''))
            for st in _cfg.special_titles
        )

        if is_toc:
            toc_section = (sec_num, section)
        elif is_special and not has_chapters:
            # 无章书：按名称判断前置/后置
            if _is_post_title(title_compact):
                post_sections.append((sec_num, section))
            else:
                pre_sections.append((sec_num, section))
        elif is_special and page_idx < first_chapter_page:
            pre_sections.append((sec_num, section))
        elif is_special and page_idx > last_chapter_page:
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

    # 输出章→节（章标题显示文字优先用目录标题，前缀按顺序重新分配）
    for ch_idx, ch_key in enumerate(chapter_order):
        if ch_key not in sections_by_chapter:
            continue
        cn = to_cn_num(ch_idx + 1)
        display = chapter_titles[ch_idx] if ch_idx < len(chapter_titles) else ch_key
        ws.append([row_num, f'{cn} {display}', 2, None])
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
def _process_directory(dir_path, special_titles=None):
    """批量处理目录下所有 *.pdf，逐个输出到 <目录>/<pdf文件名> 子目录"""
    pdfs = sorted(dir_path.glob('*.pdf'))
    if not pdfs:
        print(f'目录下未找到 PDF 文件: {dir_path}')
        return
    print(f'批量处理: {dir_path}（共 {len(pdfs)} 个PDF）')
    ok = 0
    for idx, pdf in enumerate(pdfs, 1):
        out_dir = dir_path / pdf.stem
        print(f'\n===== [{idx}/{len(pdfs)}] {pdf.name} → {out_dir} =====')
        try:
            process_pdf(pdf, out_dir, special_titles=special_titles)
            ok += 1
        except Exception as e:
            print(f'  [错误] 处理 {pdf.name} 失败: {e}')
    print(f'\n批量处理完成: 成功 {ok}/{len(pdfs)}')


def main():
    if len(sys.argv) < 2:
        print('用法:')
        print('  单文件: python extract_book.py <pdf路径> [输出目录] [特殊标题1,特殊标题2,...]')
        print('  批量:   python extract_book.py <目录> [特殊标题1,特殊标题2,...]')
        print('示例: python extract_book.py NH0633.pdf 柔石画传 "龙华英烈画传系列丛书编委会,柔石大事年表"')
        print('示例: python extract_book.py ./books')
        sys.exit(1)

    first = Path(sys.argv[1])

    # 批量模式：第一个参数为目录
    if first.is_dir():
        special_titles = None
        if len(sys.argv) >= 3:
            special_titles = [s.strip() for s in sys.argv[2].split(',') if s.strip()]
        _process_directory(first, special_titles)
        return

    # 单文件模式（保持兼容）
    pdf_path = sys.argv[1]
    if len(sys.argv) >= 3:
        output_dir = sys.argv[2]
    else:
        output_dir = first.stem

    # 可选：额外特殊标题（逗号分隔）
    special_titles = None
    if len(sys.argv) >= 4:
        special_titles = [s.strip() for s in sys.argv[3].split(',') if s.strip()]

    print(f'PDF: {pdf_path}')
    print(f'输出目录: {output_dir}')
    if special_titles:
        print(f'额外特殊标题: {special_titles}')

    process_pdf(pdf_path, output_dir, special_titles=special_titles)


if __name__ == '__main__':
    main()