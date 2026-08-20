#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
小说分段助手 · 分段与校对引擎

职责分工：
    语义判断（错别字、逗号该不该升级成句号）由 AI 在上游完成，产出「修订稿」；
    本脚本只做确定性的工作——切句、合并、编号、比对、出报表。
    这样做的好处是：长文本不会因为 AI 逐字重排而丢字，排版结果永远可复现。

输入：
    原文.txt    —— 用户粘贴的原始文本，一字未改
    修订稿.txt  —— AI 改完标点和错别字的版本（可省略，省略则只按规则分段）

输出：
    分段正文.txt   纯净正文，无序号，直接复制进作家助手
    校对报告.html  带段号的高亮对照视图，段号复制不走
    修改清单.csv   Excel 可打开的逐条修改记录

依赖：仅 Python 标准库，无需 pip 安装任何东西。
"""

import argparse
import csv
import difflib
import html
import os
import re
import sys
from datetime import datetime

# ---------------------------------------------------------------- 常量定义

# 句末标点：遇到这些字符就认为一句话说完了
SENTENCE_END = '。！？!?…'

# 收尾符号：句末标点后面如果紧跟这些，要一起归到本句末尾
CLOSING_MARKS = '”"』」）)〉》'

# 开场引号：以这些字符开头的句子视为对话，强制独立成段
DIALOG_START = '“"『「'

# 逗号类符号，用于统计单段的停顿次数
COMMA_MARKS = '，,、；;'


# ---------------------------------------------------------------- 基础工具

def count_chars(text):
    """统计有效字数：只算非空白字符（含标点，因为标点也占阅读节奏）"""
    return len([c for c in text if not c.isspace()])


def count_comma(text):
    """统计一段话里的停顿符号数量，用来判断这段是不是长句"""
    return sum(1 for c in text if c in COMMA_MARKS)


def is_dialog(text):
    """判断是不是对话句：以开场引号打头"""
    return text[:1] in DIALOG_START


def strip_punct(text):
    """
    提纯：去掉所有标点和空白，只留文字。
    用于自检——判断 AI 改稿时有没有偷偷增删内容。
    """
    return re.sub(r'[\s\W_]+', '', text, flags=re.UNICODE)


def read_text(path):
    """读文件，自动兼容 UTF-8 带签名和 GBK 两种常见编码"""
    for encoding in ('utf-8-sig', 'utf-8', 'gbk'):
        try:
            with open(path, 'r', encoding=encoding) as f:
                # 统一换行符，避免 Windows 的 \r\n 干扰偏移量计算
                return f.read().replace('\r\n', '\n').replace('\r', '\n')
        except UnicodeDecodeError:
            continue
    raise SystemExit('无法识别文件编码：%s' % path)


# ---------------------------------------------------------------- 切分逻辑

def split_blocks(full_text):
    """
    切自然段。
    用户原文里已有的换行是「大段边界」，分段只在大段内部进行，不跨行合并。
    返回 [(段落文本, 该段在全文中的起始偏移), ...]
    """
    blocks = []
    cursor = 0
    for line in full_text.split('\n'):
        stripped = line.strip()
        if stripped:
            # 记录去掉前导空白（比如首行缩进的两个全角空格）之后的真实偏移
            lead = len(line) - len(line.lstrip())
            blocks.append((stripped, cursor + lead))
        cursor += len(line) + 1  # +1 是被 split 掉的换行符
    return blocks


def split_sentences(block_text, base_offset):
    """
    把一个自然段切成句子。
    以句末标点为界，连续的句末符（比如「……」「？！」）算作一个整体，
    后面紧跟的收尾引号也归入本句。
    返回 [(句子文本, 起始偏移, 结束偏移), ...]，偏移相对于全文。
    """
    sentences = []
    start = 0
    i = 0
    n = len(block_text)

    while i < n:
        if block_text[i] in SENTENCE_END:
            end = i
            # 吞掉连续的句末标点，例如省略号由两个「…」组成
            while end + 1 < n and block_text[end + 1] in SENTENCE_END:
                end += 1
            # 吞掉收尾引号，保证「他说完了。」这样的对话完整闭合
            while end + 1 < n and block_text[end + 1] in CLOSING_MARKS:
                end += 1
            _append_sentence(sentences, block_text, start, end + 1, base_offset)
            i = end + 1
            start = i
        else:
            i += 1

    # 收尾：最后一句可能没有句末标点（作者忘了打，或者是残句）
    if start < n:
        _append_sentence(sentences, block_text, start, n, base_offset)

    return sentences


def _append_sentence(bucket, block_text, start, end, base_offset):
    """把一句话装进结果列表，同时把首尾空白从偏移里剔除，保证偏移精确对应文字"""
    piece = block_text[start:end]
    if not piece.strip():
        return
    lead = len(piece) - len(piece.lstrip())
    tail = len(piece) - len(piece.rstrip())
    bucket.append((
        piece.strip(),
        base_offset + start + lead,
        base_offset + end - tail,
    ))


def merge_sentences(sentences, opt):
    """
    把句子合并成最终段落。

    核心节奏：默认一句一段，这是网文的基本盘。
    只有在「当前段特别短」且「合并后仍然不长」时才合并，避免出现
    「他愣住了。」这种孤零零两三个字占一整段的碎片感。

    对话永远独立成段，不跟叙述混排。

    返回 [(起始偏移, 结束偏移), ...]
    """
    paragraphs = []
    current = None  # [起, 止]

    for text, start, end in sentences:
        if current is None:
            current = [start, end]
            continue

        current_text = opt['_source'][current[0]:current[1]]
        merged_len = count_chars(current_text) + count_chars(text)

        can_merge = (
            not is_dialog(text)
            and not is_dialog(current_text)
            and count_chars(current_text) <= opt['short_max']      # 只有短段才考虑接
            and merged_len <= opt['merge_max']                     # 接完不能超上限
            and count_comma(current_text + text) <= opt['max_comma']
        )

        if can_merge:
            current[1] = end
        else:
            paragraphs.append(tuple(current))
            current = [start, end]

    if current is not None:
        paragraphs.append(tuple(current))

    return paragraphs


def build_paragraphs(source_text, opt):
    """完整分段流程：自然段 → 句子 → 段落。返回段落偏移列表"""
    opt['_source'] = source_text
    result = []
    for block_text, base in split_blocks(source_text):
        sentences = split_sentences(block_text, base)
        result.extend(merge_sentences(sentences, opt))
    return result


def find_long_paragraphs(source_text, paragraphs, opt):
    """
    挑出仍然过长的段落。
    这类段落说明上游的标点升级没做到位，需要 AI 回炉再判断一次，
    脚本不擅自替用户改写内容，只负责点名。
    返回 [(段号, 字数, 逗号数, 段落文本), ...]
    """
    flagged = []
    for index, (start, end) in enumerate(paragraphs, 1):
        text = source_text[start:end]
        chars = count_chars(text)
        commas = count_comma(text)
        if chars > opt['para_max'] or commas > opt['max_comma']:
            flagged.append((index, chars, commas, text))
    return flagged


# ---------------------------------------------------------------- 差异比对

def diff_changes(origin, revised):
    """
    逐字比对原文和修订稿，找出所有改动。

    difflib 会产出很多零碎的单字差异，这里做一次归并：
    间距很近的改动合成一条，读起来才像人话，而不是一堆单字条目。

    返回 [{'原文':…, '改后':…, 'j1':…, 'j2':…}, ...]，j 是在修订稿中的位置。
    """
    matcher = difflib.SequenceMatcher(None, origin, revised, autojunk=False)
    raw = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            continue
        before, after = origin[i1:i2], revised[j1:j2]
        # 纯换行、纯空白的差异属于排版噪音，不算修改
        if not before.strip() and not after.strip():
            continue
        raw.append({'原文': before, '改后': after, 'i1': i1, 'i2': i2, 'j1': j1, 'j2': j2})

    # 归并相邻改动：中间夹着不超过 2 个字符时视为同一处修改
    merged = []
    for item in raw:
        if merged:
            prev = merged[-1]
            gap_new = item['j1'] - prev['j2']
            gap_old = item['i1'] - prev['i2']
            if 0 <= gap_new <= 2 and 0 <= gap_old <= 2:
                prev['原文'] += origin[prev['i2']:item['i1']] + item['原文']
                prev['改后'] += revised[prev['j2']:item['j1']] + item['改后']
                prev['i2'], prev['j2'] = item['i2'], item['j2']
                continue
        merged.append(dict(item))

    return merged


def classify(before, after):
    """给修改归类，方便用户在表格里筛选"""
    if not after.strip():
        return '删除'
    if not before.strip():
        return '补字'
    if not strip_punct(before) and not strip_punct(after):
        return '标点'
    if len(before) == len(after):
        return '错别字'
    return '改写'


def locate(changes, paragraphs, source_text):
    """
    把每处修改定位到「第几段、第几句」。
    段号严格等于分段正文里的段落顺序，这就是用户要的那根坐标轴。
    """
    for change in changes:
        change['段号'] = 0
        change['句序'] = 0
        for index, (start, end) in enumerate(paragraphs, 1):
            if start <= change['j1'] < end or (start <= change['j1'] <= end and change['j1'] == change['j2']):
                change['段号'] = index
                # 段内第几句：数一下改动位置之前出现过几个句末标点
                head = source_text[start:change['j1']]
                change['句序'] = sum(1 for c in head if c in SENTENCE_END) + 1
                break
    return changes


def self_check(origin, revised):
    """
    防幻觉自检。
    AI 改稿最大的风险是「顺手润色」，把用户的原句悄悄改写了。
    这里剥掉全部标点只比文字，一旦相似度掉下来就报警。
    返回警告文案列表，空列表表示体检通过。
    """
    warnings = []
    pure_origin, pure_revised = strip_punct(origin), strip_punct(revised)

    ratio = difflib.SequenceMatcher(None, pure_origin, pure_revised, autojunk=False).ratio()
    delta = len(pure_revised) - len(pure_origin)

    if ratio < 0.98:
        warnings.append(
            '文字相似度仅 %.2f%%，修订稿可能改动了内容而不只是标点和错字，请复核。' % (ratio * 100)
        )
    if abs(delta) > max(5, len(pure_origin) * 0.01):
        warnings.append(
            '纯文字净%s %d 字（原文 %d 字 → 修订稿 %d 字），疑似增删内容。'
            % ('增' if delta > 0 else '减', abs(delta), len(pure_origin), len(pure_revised))
        )

    origin_blocks, revised_blocks = len(split_blocks(origin)), len(split_blocks(revised))
    if origin_blocks != revised_blocks:
        warnings.append(
            '自然段数量不一致（原文 %d 段 / 修订稿 %d 段），修订稿可能漏处理或多切了段落。'
            % (origin_blocks, revised_blocks)
        )

    return warnings


# ---------------------------------------------------------------- 一致性检查

# 高频虚词。人名里基本不会出现这些字，用来把噪音组合筛掉
FUNCTION_CHARS = set(
    '的了是在和就都也不这那什么个上下中有他她它我你们之其而于与把被给对从到说道着过来去'
    '很再又还只将且但如若因所需要能会没最多少大小时候现如今算样着起出向前后里外边面'
    '一二三四五六七八九十百千万些几每各同样种样儿子头手眼心目声色气力天地人'
)

# 常见双字词语，长得像人名但不是（如「明白」「事情」），从候选里剔掉降低误报
COMMON_WORDS = {
    '明白', '事情', '时候', '知道', '觉得', '感觉', '朋友', '自己', '什么', '因为',
    '所以', '如果', '已经', '没有', '可以', '这样', '那样', '怎么', '哪里', '这里',
    '那里', '这些', '那些', '这个', '那个', '我们', '你们', '他们', '她们', '它们',
    '起来', '出来', '过来', '回去', '一起', '一直', '一定', '一样', '一些', '有的',
    '的话', '一种', '一般', '对于', '关于', '通过', '由于', '并且', '而且', '或者',
    '然而', '片刻', '之间', '突然', '忽然', '渐渐', '终于', '毕竟', '显然', '自然',
    '当然', '原来', '究竟', '实在', '看来',
}


def looks_like_name(gram):
    """粗判一个字符组合像不像专有名词：不含高频虚词、且不是常见词，才有嫌疑"""
    if gram in COMMON_WORDS:
        return False
    return not any(c in FUNCTION_CHARS for c in gram)


def count_ngrams(text, size):
    """统计指定长度的连续汉字组合出现次数"""
    counter = {}
    for block in re.findall(r'[\u4e00-\u9fff]{%d,}' % size, text):
        for i in range(len(block) - size + 1):
            gram = block[i:i + size]
            counter[gram] = counter.get(gram, 0) + 1
    return counter


def detect_name_typo(text, main_min=2, ratio=2.0, pool=800):
    """
    专有名词一致性检测。

    网文作者最常犯的错之一：同一个角色写着写着名字少一笔，或换成同音字。
    这类错误人眼极难发现，但对读者非常出戏。

    判定思路：统计 2~4 字的汉字组合，找出「长度相同、只差一个字、
    且一方明显更高频」的配对——高频的多半是正名，低频的多半是笔误。

    只报告，绝不自动改：也可能真是两个角色（比如同族兄弟）。
    返回 [(正名, 正名频次, 疑似, 疑似频次), ...]
    """
    buckets = {size: count_ngrams(text, size) for size in (2, 3, 4)}

    # 先剔除冗余的短组合：如果「明博」永远只出现在「周明博」里（两者频次相同），
    # 那它没有独立信息量，留着只会制造「明博 / 明白」这类误报。
    # 长度 L+1 的组合只包含两个长度 L 的子串，据此可以一遍扫完，不必两两比对。
    redundant = set()
    for size in (2, 3):
        longer = buckets[size + 1]
        shorter = buckets[size]
        for gram, count in longer.items():
            # 只对够得上「正名」资格的高频组合去冗余。
            # 低频组合的超串必然同样低频，一刀切会把真正的笔误（往往只出现一两次）误杀。
            if count < main_min:
                continue
            for sub in (gram[:-1], gram[1:]):
                if shorter.get(sub) == count:
                    redundant.add(sub)

    suspects = []
    for size in (2, 3, 4):
        candidates = [
            (g, c) for g, c in buckets[size].items()
            if looks_like_name(g) and g not in redundant
        ]
        candidates.sort(key=lambda item: -item[1])
        candidates = candidates[:pool]

        for i, (gram_a, count_a) in enumerate(candidates):
            if count_a < main_min:
                continue
            for gram_b, count_b in candidates[i + 1:]:
                if count_b >= count_a:
                    continue
                # 高频方必须明显占优，否则可能真是两个不同的角色
                if count_a < count_b * ratio:
                    continue
                # 只差一个字才算笔误
                if sum(1 for x, y in zip(gram_a, gram_b) if x != y) == 1:
                    suspects.append((gram_a, count_a, gram_b, count_b))

    # 同一对名字可能在不同长度上重复命中，保留信息量最大的那条
    suspects.sort(key=lambda item: (-len(item[0]), -item[1]))
    kept = []
    for item in suspects:
        if any(item[0] in seen[0] and item[2] in seen[2] for seen in kept):
            continue
        kept.append(item)
    return kept[:20]


def check_quotes(text):
    """
    用「引号深度」判断引号是否成对，而不是简单数比多少个。

    中文跨段对话是「每段都开前引号、只在末尾闭后引号」，
    所以全文前引号数量本就会多于后引号——直接数比多少是错的。
    正确做法：像括号匹配一样走一遍，最后深度为 0 才算闭合。
    """
    warnings = []
    depth = 0
    bad_lines = []
    for no, line in enumerate(text.split('\n'), 1):
        for ch in line:
            if ch == '“':
                depth += 1
            elif ch == '”':
                depth -= 1
        if depth < 0:
            bad_lines.append(no)
    if depth != 0:
        warnings.append(
            '引号未成对：全文走完仍有 %d 层引号没闭合，可能漏打了后引号。' % depth
        )
    if bad_lines:
        warnings.append('第 %s 行出现「后引号多于前引号」，请检查是否多打了引号。'
                        % '、'.join(str(n) for n in bad_lines[:8]))
    return warnings


# ---------------------------------------------------------------- 标点与排版增强

# 半角标点 → 全角标点的映射（引号单独处理，因为它要成对定向）
FWHW_MAP = {
    ',': '，', '.': '。', '!': '！', '?': '？', ';': '；', ':': '：',
    '(': '（', ')': '）', '[': '【', ']': '】',
}


def _is_cjk(ch):
    """判断一个字符是不是中文字或中文标点，用于决定半角标点该不该转全角"""
    return ('\u4e00' <= ch <= '\u9fff') or ch in '，。！？、；：“”‘’（）【】《》…—'


def normalize_fwhw(text):
    """
    半角标点审查 → 转全角。

    OCR、手机输入、从网页复制常常混入半角逗号句号。
    但数字里的半角点（如 3.14）不能动，所以只在该标点旁边有中文字时才转换，
    避免破坏英文和数字。半角双引号成对出现，按顺序转成前/后引号。
    """
    out = []
    q_parity = 0
    n = len(text)
    for i, ch in enumerate(text):
        if ch == '"':
            out.append('“' if q_parity % 2 == 0 else '”')
            q_parity += 1
            continue
        if ch in FWHW_MAP:
            prev = text[i - 1] if i > 0 else ''
            nxt = text[i + 1] if i + 1 < n else ''
            if _is_cjk(prev) or _is_cjk(nxt):
                out.append(FWHW_MAP[ch])
                continue
        out.append(ch)
    return ''.join(out)


def rebalance_quotes(para_texts):
    """
    引号处理，解决作者两个习惯问题：

    1. 多段对话只首尾加引号 —— 按中文规范，跨段引号每段开头都要有前引号，
       只有最后一段结尾有后引号。本函数给中段补上前引号。
    2. 引号残缺 —— 走完一遍如果仍处在「引号内」，说明文末缺后引号，补上。

    返回 (修正后的段落列表, 改动说明列表)。
    """
    corrected = []
    notes = []
    depth = 0  # 当前位于第几层引号之内
    for idx, p in enumerate(para_texts):
        stripped = p.lstrip()
        if depth > 0 and not stripped.startswith('“'):
            p = '“' + p
            notes.append((idx + 1, '为多段对话补上段首前引号'))
        for ch in p:
            if ch == '“':
                depth += 1
            elif ch == '”':
                depth -= 1
        corrected.append(p)
    if depth > 0:
        corrected[-1] = corrected[-1] + '”'
        notes.append((len(corrected), '文末补全残缺的后引号'))
    elif depth < 0:
        notes.append((0, '发现多余的后引号（前引号少于后引号），请检查'))
    return corrected, notes


# 常见错字短语表（仅做整词替换，不做单字猜测，降低误报）
COMMON_TYPO = {
    '帐号': '账号', '按装': '安装', '布署': '部署', '沉缅': '沉湎',
    '防碍': '妨碍', '辐度': '幅度', '观摹': '观摩', '寒喧': '寒暄',
    '穿流不息': '川流不息', '璀灿': '璀璨', '打腊': '打蜡',
    '凋蔽': '凋敝', '钉书机': '订书机', '渡假': '度假',
    '烦燥': '烦躁', '鼓午': '鼓舞', '涵概': '涵盖', '羁拌': '羁绊',
    '简炼': '简练', '精萃': '精粹', '肯求': '恳求', '蜡梅': '腊梅',
    '老俩口': '老两口', '了草': '潦草', '灵柩': '灵柩', '荧火虫': '萤火虫',
    '萎糜': '萎靡', '偏辟': '偏僻', '拼揍': '拼凑', '善长': '擅长',
    '松驰': '松弛', '坦护': '袒护', '惦量': '掂量', '余辉': '余晖',
    '急燥': '急躁', '弦律': '旋律', '一柱香': '一炷香', '涨溢': '洋溢',
}


def fix_common_typo(text):
    """整词替换常见错字，返回 (修正文本, 改动说明)。仅在显式开启时使用。"""
    notes = []
    for wrong, right in COMMON_TYPO.items():
        if wrong in text:
            count = text.count(wrong)
            text = text.replace(wrong, right)
            notes.append('「%s」→「%s」×%d' % (wrong, right, count))
    return text, notes


# ---------------------------------------------------------------- 标题生成

PLOT_WORDS = set(
    '说答话讲问回喊叫看看望注视盯瞪听笑哭想想明知道觉猜估判决走跑冲扑'
    '抓拿举抬转回退闪躲攻防杀战斗打摔撞怕惊慌怒喜乐爱恨怨仇疑惑悟醒悔'
    '羞怯沉默瞧瞥奔逃迎送递捧托扛背骑坐站躺冲出闯进逼退闪过'
)


def top_name(text):
    """从全文挑出最可能的角色名：先用人名笔误检测的正名，再退而求其次取高频名词组合"""
    suspects = detect_name_typo(text)
    if suspects:
        return suspects[0][0]
    for size in (3, 2, 4):
        counter = count_ngrams(text, size)
        cand = [(g, c) for g, c in counter.items() if looks_like_name(g)]
        cand.sort(key=lambda item: -item[1])
        if cand:
            return cand[0][0]
    return ''


def plot_score(sentence):
    """给句子打「剧情含量」分，分高的更适合作标题素材"""
    s = strip_punct(sentence)
    score = sum(1 for c in s if c in PLOT_WORDS)
    if s.startswith('“'):
        score += 2
    return score


def _best_sentence(sentences):
    if not sentences:
        return ''
    return max(sentences, key=lambda s: plot_score(s) - len(strip_punct(s)) // 20)


# 适合作为标题断句点的字（在其后断开读起来自然）
BOUNDARY = set(
    '了得的着过在与向对把被给及和而到如若因看说问道走拿举抬转回退闪躲'
    '攻防杀战打想明知觉猜决冲闯迎送闪现露露出睁'
)


def _snap_cut(text, limit):
    """截到 limit 个汉字，但尽量落在断句点之后，避免把词从中间切断"""
    if len(text) <= limit:
        return text
    pos = limit
    for i in range(min(limit, len(text)), 0, -1):
        if text[i - 1] in BOUNDARY:
            pos = i
            break
    return text[:pos]


def _pick_phrase(sentences, limit):
    """从句子里挑一句：优先挑本身就不超字数的一句（整句更自然），否则截最短的一句"""
    cleaned = [(s, strip_punct(s)) for s in sentences if strip_punct(s).strip()]
    fit = [c for _, c in cleaned if 0 < len(c) <= limit]
    if fit:
        fit.sort(key=lambda c: -plot_score(c))
        return fit[0]
    if cleaned:
        shortest = min(cleaned, key=lambda item: len(item[1]))[1]
        return _snap_cut(shortest, limit)
    return ''


def _trim_chars(text, limit):
    """只按非空格字符计数截断，保证标题不超过 limit 个汉字"""
    out, count = [], 0
    for ch in text:
        if ch == ' ':
            out.append(ch)
            continue
        if count >= limit:
            break
        out.append(ch)
        count += 1
    return ''.join(out)


def generate_title(text, fmt='自动', limit=10):
    """
    启发式总结标题（无 AI 时的兜底方案，网页版也用它）。

    - 简明：挑一句本身就不超字数的话作标题，最自然；挑不到再截断
    - 章回体：前后两半各挑一句拼成「XXXX XXXX」
    - 自动：默认走简明，最稳妥

    硬上限 limit（默认 10）个字，超过一定截断。
    """
    sentences = []
    for block, base in split_blocks(text):
        for s, _, _ in split_sentences(block, 0):
            if s.strip():
                sentences.append(s.strip())
    if not sentences:
        sentences = [text]

    if fmt == '章回体':
        half = max(1, len(sentences) // 2)
        p1 = _pick_phrase(sentences[:half], limit // 2) or _snap_cut(strip_punct(sentences[0]), limit // 2)
        p2 = _pick_phrase(sentences[half:], limit // 2) or _snap_cut(strip_punct(sentences[-1]), limit // 2)
        title = (p1 + ' ' + p2).strip()
    else:  # 简明 / 自动
        title = _pick_phrase(sentences, limit)

    if not title:
        title = strip_punct(text)[:limit] or '未命名'
    if fmt != '章回体':
        title = _trim_chars(title, limit)
    return title


def clip_context(text, start, end, pad=4):
    """截取改动片段及其前后若干字，让表格里的记录能看懂"""
    left, right = max(0, start - pad), min(len(text), end + pad)
    prefix = '…' if left > 0 else ''
    suffix = '…' if right < len(text) else ''
    return (prefix + text[left:right] + suffix).replace('\n', ' ')


# ---------------------------------------------------------------- 产物输出

def write_plain(path, source_text, paragraphs, title=''):
    """
    输出纯净正文。
    段首带一行标题（可单独删），段落之间空一行，不带任何序号、标记、说明文字——
    用户全选复制粘贴进作家助手，格式就是对的。
    """
    body = '\n\n'.join(source_text[s:e] for s, e in paragraphs)
    if title:
        body = title + '\n\n' + body
    with open(path, 'w', encoding='utf-8') as f:
        f.write(body + '\n')
    return body


def write_csv(path, changes, long_paragraphs, name_suspects, title=''):
    """
    输出修改清单。
    带 UTF-8 BOM，Excel 双击打开不会乱码。
    「说明」列留空，交给 AI 在上游补写修改理由。
    """
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        if title:
            writer.writerow(['【标题】', title])
            writer.writerow([])
        writer.writerow(['序号', '段号', '段内第几句', '类型', '原文', '修改后', '前后对照', '说明'])
        for number, change in enumerate(changes, 1):
            writer.writerow([
                number,
                change.get('段号', ''),
                change.get('句序', ''),
                classify(change['原文'], change['改后']),
                change['原文'],
                change['改后'],
                change.get('对照', ''),
                '',
            ])

        if long_paragraphs:
            writer.writerow([])
            writer.writerow(['【偏长段落】以下段落仍然偏长，建议回到修订稿再判断一次标点'])
            writer.writerow(['段号', '字数', '逗号数', '段落内容'])
            for index, chars, commas, text in long_paragraphs:
                writer.writerow([index, chars, commas, text])

        if name_suspects:
            writer.writerow([])
            writer.writerow(['【待确认】疑似专有名词不一致，脚本不会自动改，请作者自行判断'])
            writer.writerow(['疑似正名', '出现次数', '疑似笔误', '出现次数'])
            for main_name, main_count, typo, typo_count in name_suspects:
                writer.writerow([main_name, main_count, typo, typo_count])


HTML_TEMPLATE = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>校对报告 · 小说分段助手</title>
<style>
  :root {{
    --bg: #f7f7f5; --card: #ffffff; --ink: #1f2328; --muted: #8b8b86;
    --line: #e6e6e1; --mark: #fff3b0; --mark-line: #e8c547; --del: #ffd9d4;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0; padding: 32px 20px 64px; background: var(--bg); color: var(--ink);
    font-family: "PingFang SC", "Microsoft YaHei", "Source Han Sans SC", sans-serif;
    line-height: 1.9;
  }}
  .wrap {{ max-width: 860px; margin: 0 auto; }}
  h1 {{ font-size: 20px; margin: 0 0 4px; letter-spacing: .5px; }}
  .sub {{ color: var(--muted); font-size: 13px; margin-bottom: 24px; }}
  .stats {{
    display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 20px;
  }}
  .chip {{
    background: var(--card); border: 1px solid var(--line); border-radius: 999px;
    padding: 5px 14px; font-size: 13px; color: #4a4a44;
  }}
  .chip b {{ color: var(--ink); font-weight: 600; }}
  .warn {{
    background: #fff6e5; border: 1px solid #f0d49a; border-radius: 8px;
    padding: 12px 16px; font-size: 13.5px; margin-bottom: 20px; color: #7a5a12;
  }}
  .warn p {{ margin: 4px 0; }}
  .bar {{ display: flex; gap: 10px; margin-bottom: 16px; }}
  button {{
    border: 1px solid var(--line); background: var(--card); color: var(--ink);
    border-radius: 8px; padding: 8px 16px; font-size: 13.5px; cursor: pointer;
    font-family: inherit;
  }}
  button:hover {{ background: #efefe9; }}
  .paper {{
    background: var(--card); border: 1px solid var(--line); border-radius: 12px;
    padding: 28px 32px 28px 20px; counter-reset: para;
  }}
  .para {{
    counter-increment: para; position: relative; padding-left: 3.4em;
    margin: 0 0 1.15em; font-size: 16px;
  }}
  /* 段号用伪元素生成：只看得见，复制文字时不会被带走 */
  .para::before {{
    content: counter(para);
    position: absolute; left: 0; top: 0; width: 2.4em;
    text-align: right; color: #c2c2ba; font-size: 12.5px; line-height: 2.4;
    user-select: none; -webkit-user-select: none; -moz-user-select: none;
  }}
  .para:hover::before {{ color: #8a8a80; }}
  mark {{
    background: var(--mark); border-bottom: 1.5px solid var(--mark-line);
    padding: 1px 2px; border-radius: 2px; cursor: help;
  }}
  mark.del {{ background: var(--del); border-bottom-color: #e08d80; padding: 1px 4px; }}
  .long {{ background: #f4f7ff; box-shadow: inset 3px 0 0 #9db4ff; border-radius: 4px; }}
  .foot {{ color: var(--muted); font-size: 12.5px; margin-top: 24px; text-align: center; }}
  textarea {{ position: absolute; left: -9999px; }}
  @media print {{ body {{ background: #fff; }} .bar {{ display: none; }} }}
</style>
</head>
<body>
<div class="wrap">
  <h1>{title}</h1>
  <div class="sub">{timestamp} · 黄底为改动处，鼠标悬停可看原文；左侧段号仅用于对照，复制文字时不会带上</div>

  <div class="stats">
    <span class="chip">原文 <b>{origin_chars}</b> 字</span>
    <span class="chip">分成 <b>{para_count}</b> 段</span>
    <span class="chip">平均每段 <b>{avg_chars}</b> 字</span>
    <span class="chip">共修改 <b>{change_count}</b> 处</span>
    <span class="chip">偏长段落 <b>{long_count}</b> 段</span>
  </div>

  {warn_block}

  <div class="bar">
    <button onclick="copyClean(this)">复制纯净正文</button>
    <button onclick="window.print()">打印 / 存为 PDF</button>
  </div>

  <div class="paper">
{paragraphs}
  </div>

  <div class="foot">小说分段助手 · 段号是坐标轴，复制走的永远是干净正文</div>
</div>

<textarea id="clean">{clean_text}</textarea>
<script>
function copyClean(btn) {{
  var box = document.getElementById('clean');
  var done = function () {{
    btn.textContent = '已复制 ✓';
    setTimeout(function () {{ btn.textContent = '复制纯净正文'; }}, 1600);
  }};
  // 优先用现代剪贴板接口，本地打开文件时降级到旧接口
  if (navigator.clipboard && window.isSecureContext) {{
    navigator.clipboard.writeText(box.value).then(done, function () {{ legacy(); }});
  }} else {{
    legacy();
  }}
  function legacy() {{
    box.style.position = 'static';
    box.select();
    try {{ document.execCommand('copy'); done(); }}
    catch (e) {{ alert('复制失败，请手动全选正文'); }}
    box.style.position = 'absolute';
  }}
}}
</script>
</body>
</html>
'''


def write_html(path, source_text, paragraphs, changes, long_indexes,
               warnings, origin_chars, name_suspects=(), title=''):
    """
    输出校对报告网页。

    两个关键设计：
    1. 段号用 CSS counter + 伪元素生成，用户复制时序号不会混进正文；
    2. 改动处用 <mark> 高亮，title 属性挂原文，悬停即可对照。
    """
    # 按段落归拢改动，便于在段内定位高亮区间
    by_para = {}
    for change in changes:
        by_para.setdefault(change.get('段号', 0), []).append(change)

    blocks = []
    for index, (start, end) in enumerate(paragraphs, 1):
        marks = sorted(by_para.get(index, []), key=lambda c: c['j1'])
        pieces = []
        cursor = start
        for change in marks:
            j1 = max(change['j1'], start)
            j2 = min(change['j2'], end)
            if j1 < cursor:
                continue
            pieces.append(html.escape(source_text[cursor:j1]))
            tip = html.escape('原文：' + (change['原文'].strip() or '（无，此处为新增）'), quote=True)
            if j2 > j1:
                pieces.append('<mark title="%s">%s</mark>' % (tip, html.escape(source_text[j1:j2])))
            else:
                # 纯删除：修订稿里没有可高亮的字，用一个红点标出位置
                pieces.append('<mark class="del" title="%s">·</mark>' % tip)
            cursor = max(cursor, j2)
        pieces.append(html.escape(source_text[cursor:end]))

        css_class = 'para long' if index in long_indexes else 'para'
        blocks.append('    <p class="%s">%s</p>' % (css_class, ''.join(pieces)))

    # 警告区：体检问题 + 疑似人名笔误，都只提示不代劳
    warn_items = ['⚠ ' + w for w in warnings]
    for main_name, main_count, typo, typo_count in name_suspects:
        warn_items.append(
            '❓ 疑似专有名词不一致：「%s」出现 %d 次，「%s」出现 %d 次，请确认是否笔误'
            % (main_name, main_count, typo, typo_count)
        )
    warn_block = ''
    if warn_items:
        items = ''.join('<p>%s</p>' % html.escape(item) for item in warn_items)
        warn_block = '<div class="warn">%s</div>' % items

    clean_text = '\n\n'.join(source_text[s:e] for s, e in paragraphs)
    total = sum(count_chars(source_text[s:e]) for s, e in paragraphs)

    page = HTML_TEMPLATE.format(
        timestamp=datetime.now().strftime('%Y-%m-%d %H:%M'),
        title=html.escape(title) or '校对报告',
        origin_chars=origin_chars,
        para_count=len(paragraphs),
        avg_chars=(total // len(paragraphs)) if paragraphs else 0,
        change_count=len(changes),
        long_count=len(long_indexes),
        warn_block=warn_block,
        paragraphs='\n'.join(blocks),
        clean_text=html.escape(clean_text),
    )
    with open(path, 'w', encoding='utf-8') as f:
        f.write(page)


# ---------------------------------------------------------------- 主流程

def build_parser():
    """命令行参数：中文名和英文名都能用，方便不同 Agent 应用调用"""
    parser = argparse.ArgumentParser(
        prog='fenduan.py',
        description='小说分段助手 · 把长段落切成网文节奏的短段，并输出校对报告',
    )
    parser.add_argument('--原文', '--yuanwen', dest='origin', required=True,
                        help='原始文本文件路径（必填）')
    parser.add_argument('--修订稿', '--xiuding', dest='revised', default=None,
                        help='AI 改完标点错字的文本路径；不给则只按规则分段')
    parser.add_argument('--输出目录', '--outdir', dest='outdir', default='.',
                        help='三份产物的存放目录，默认当前目录')
    parser.add_argument('--合并上限', '--merge-max', dest='merge_max', type=int, default=24,
                        help='两个短句合并后的字数上限，调大段落变长（默认 24）')
    parser.add_argument('--段落上限', '--para-max', dest='para_max', type=int, default=45,
                        help='单段字数硬上限，超出会被点名（默认 45）')
    parser.add_argument('--最多逗号', '--max-comma', dest='max_comma', type=int, default=2,
                        help='单段允许的逗号数量（默认 2）')
    parser.add_argument('--短句阈值', '--short-max', dest='short_max', type=int, default=6,
                        help='短于这个字数的段落才考虑向后合并，调大→段落变少（默认 6）')
    parser.add_argument('--标题', '--title', dest='title', default='',
                        help='AI 给定的标题；不填则由脚本启发式生成（网页版同款逻辑）')
    parser.add_argument('--标题格式', '--title-fmt', dest='title_fmt', default='自动',
                        choices=['简明', '章回体', '自动'],
                        help='标题风格：简明 / 章回体 / 自动（默认）')
    parser.add_argument('--标题上限', '--title-limit', dest='title_limit', type=int, default=10,
                        help='标题最多字数，硬上限（默认 10）')
    parser.add_argument('--半角转全角', '--fwhw', dest='fwhw', action='store_true', default=True,
                        help='半角标点转全角（默认开）')
    parser.add_argument('--不转全角', '--no-fwhw', dest='fwhw', action='store_false',
                        help='关闭半角转全角')
    parser.add_argument('--补全引号', '--fix-quotes', dest='fix_quotes', action='store_true', default=True,
                        help='补全多段对话引号与文末残缺引号（默认开）')
    parser.add_argument('--不补全引号', '--no-fix-quotes', dest='fix_quotes', action='store_false',
                        help='关闭引号补全')
    parser.add_argument('--纠常错字', '--fix-common', dest='fix_common', action='store_true', default=False,
                        help='整词替换常见错字（默认关，避免误改；网页版默认也关）')
    return parser


def main():
    # Windows 控制台默认 GBK，先切成 UTF-8，否则打印中文会报错
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except Exception:
            pass

    args = build_parser().parse_args()

    origin = read_text(args.origin)
    base = read_text(args.revised) if args.revised else origin

    # 0、排版增强：半角转全角 → 可选常见错字（都只在显式/默认开启时做）
    if args.fwhw:
        base = normalize_fwhw(base)
    common_notes = []
    if args.fix_common:
        base, common_notes = fix_common_typo(base)

    opt = {
        'merge_max': args.merge_max,
        'para_max': args.para_max,
        'max_comma': args.max_comma,
        'short_max': args.short_max,
    }

    # 一、按修订稿分段（分段结果的坐标轴以 base 为准）
    paragraphs = build_paragraphs(base, opt)

    # 二、取段落文本，做引号处理（多段对话补引号 + 文末残缺补全）
    para_texts = [base[s:e] for s, e in paragraphs]
    quote_notes = []
    if args.fix_quotes:
        para_texts, quote_notes = rebalance_quotes(para_texts)

    # 三、重组为正文，重算段落偏移，让后续高亮和清单坐标与修正后的文本对齐
    corrected = '\n\n'.join(para_texts)
    new_paragraphs = []
    cursor = 0
    for t in para_texts:
        new_paragraphs.append((cursor, cursor + len(t)))
        cursor += len(t) + 2  # +2 是段落之间的 \n\n
    source_text = corrected
    paragraphs = new_paragraphs

    # 四、标题：优先用 AI 给定的，否则启发式生成
    title = args.title if args.title else generate_title(base, args.title_fmt, args.title_limit)

    # 五、逐字比对，找出改了哪些地方，并定位到段
    changes = []
    if args.revised:
        changes = locate(diff_changes(origin, corrected), paragraphs, source_text)
        for change in changes:
            # 单个逗号这种改动光看片段说明不了问题，带上前后各几个字才读得懂
            change['对照'] = '%s  ⇒  %s' % (
                clip_context(origin, change['i1'], change['i2']),
                clip_context(source_text, change['j1'], change['j2']),
            )

    # 六、体检：确认 AI 没顺手改写内容，顺便查引号完整性
    warnings = self_check(origin, base) if args.revised else []
    warnings.extend(check_quotes(base))
    for idx, note in quote_notes:
        warnings.append('引号处理 · 第 %d 段：%s' % (idx, note))
    for note in common_notes:
        warnings.append('常见错字 · ' + note)

    # 七、点名仍然偏长的段落，以及疑似写错的人名
    long_paragraphs = find_long_paragraphs(source_text, paragraphs, opt)
    long_indexes = set(item[0] for item in long_paragraphs)
    name_suspects = detect_name_typo(source_text)

    # 八、出三份产物
    outdir = args.outdir
    os.makedirs(outdir, exist_ok=True)
    plain_path = os.path.join(outdir, '分段正文.txt')
    html_path = os.path.join(outdir, '校对报告.html')
    csv_path = os.path.join(outdir, '修改清单.csv')

    write_plain(plain_path, source_text, paragraphs, title)
    write_html(html_path, source_text, paragraphs, changes, long_indexes,
               warnings, count_chars(origin), name_suspects, title)
    write_csv(csv_path, changes, long_paragraphs, name_suspects, title)

    # 九、把关键信息打给上游，AI 据此决定要不要回炉
    print('分段完成')
    print('  标题：%s' % title)
    print('  原文字数：%d' % count_chars(origin))
    print('  段落数量：%d' % len(paragraphs))
    print('  修改处数：%d' % len(changes))
    print('  偏长段落：%d' % len(long_paragraphs))
    if long_paragraphs:
        for index, chars, commas, text in long_paragraphs[:10]:
            preview = text if len(text) <= 40 else text[:40] + '…'
            print('    第 %d 段（%d 字 / %d 个逗号）：%s' % (index, chars, commas, preview))
    if name_suspects:
        print('  待确认的专有名词：')
        for main_name, main_count, typo, typo_count in name_suspects:
            print('    ❓「%s」×%d  与「%s」×%d 只差一个字，请确认是否笔误'
                  % (main_name, main_count, typo, typo_count))
    if warnings:
        print('  体检警告：')
        for warning in warnings:
            print('    ⚠ %s' % warning)
    else:
        print('  体检结果：通过，修订稿只动了标点和错字')
    print('产物目录：%s' % os.path.abspath(outdir))


if __name__ == '__main__':
    main()
