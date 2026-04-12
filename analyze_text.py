#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
九一八事变历史语料NLP分析脚本
功能：中俄文文本分离、词频统计、词云生成、数据可视化
作者：数据分析专家
日期：2026-04-07
"""

import re
import os
from collections import Counter, defaultdict
from typing import List, Dict, Tuple

import jieba
import jieba.posseg as pseg
import pandas as pd
import matplotlib.pyplot as plt
from wordcloud import WordCloud
from docx import Document
import pymorphy3
from nltk.corpus import stopwords
import nltk
from deep_translator import GoogleTranslator

# ============================================================================
# 全局配置区域
# ============================================================================

# 中文同义词映射字典（可手动维护）
CHN_SYNONYMS = {
    '日本军': '日军',
    # '日本人': '日本',
    '中国军': '中国军队',
    '东北军': '中国军队',
    # 在此添加更多同义词映射
}

# 俄语同义词映射字典（可手动维护）
RUS_SYNONYMS = {
    'японцы': 'японский',
    'японец': 'японский',
    'китайцы': 'китайский',
    'китаец': 'китайский',
    # 在此添加更多同义词映射
}

# 俄语到中文的翻译映射字典（手动维护，优先级高于API翻译）
RU_TO_ZH_TRANS = {
    'японский': '日本',
    'китайский': '中国',
    'война': '战争',
    'армия': '军队',
    'войска': '部队',
    # 在此添加更多翻译映射，避免每次都调用API
}

# ============================================================================
# 停用词加载与配置
# ============================================================================

def load_chinese_stopwords():
    """
    加载中文停用词表
    1. 从 stopwords_cn.txt 文件读取基础停用词
    2. 追加领域专属无意义词汇（新闻报道类、抽象词）
    """
    stopwords_set = set()

    # 读取外部停用词文件
    stopwords_file = 'resources/stopwords_cn.txt'
    if os.path.exists(stopwords_file):
        with open(stopwords_file, 'r', encoding='utf-8') as f:
            for line in f:
                word = line.strip()
                if word:  # 跳过空行
                    stopwords_set.add(word)
        print(f"✓ 已加载 {len(stopwords_set)} 个基础停用词（来自 {stopwords_file}）")
    else:
        print(f"⚠ 未找到停用词文件：{stopwords_file}，使用内置停用词")

    # 追加领域专属无意义词汇（针对抗日战争历史研究）
    domain_stopwords = {
        # 新闻报道套话
        '记者', '报道', '消息', '通讯社', '声明', '写道',
        # 抽象泛义词
        '发生', '问题', '进行', '要求', '方面', '行动', '认为', '局势', '继续', '准备',
        '活动', '地区', '事件', '当局', '情况', '表示', '指出', '提出', '作出', '采取',
        '开始', '结束', '出现', '成为', '工作', '建立', '决定', '同意', '大会', '会议',
        # 时空和逻辑连接词
        '同时', '正在', '如果', '立即', '任何', '以及', '不断', '之间', '展开',
        # 代词和量词
        '多个', '他们', '我们', '这些', '那些', '一切', '一项', '一名',
    }

    stopwords_set.update(domain_stopwords)
    print(f"✓ 追加 {len(domain_stopwords)} 个领域专属停用词")
    print(f"✓ 中文停用词总数：{len(stopwords_set)}")

    return stopwords_set


def load_russian_stopwords():
    """
    加载俄语停用词表
    1. 基于 NLTK 的俄语停用词
    2. 追加新闻报道类词汇（对应中文套话）
    3. 追加顽固代词和系动词的 lemma 形式
    """
    # 基础停用词（NLTK）
    try:
        nltk.data.find('corpora/stopwords')
    except LookupError:
        nltk.download('stopwords')

    stopwords_set = set(stopwords.words('russian'))
    print(f"✓ 已加载 {len(stopwords_set)} 个 NLTK 俄语停用词")

    # 追加新闻报道类词汇（对应中文套话和抽象名词）
    news_stopwords = {
        'корреспондент',  # 记者
        'сообщение', 'сообщать', 'сообщить',  # 消息、报道
        'происходить',  # 发生
        'вопрос',  # 问题
        'проводить',  # 进行
        'требовать',  # 要求
        'сторона',  # 方面
        'действие',  # 行动
        'считать',  # 认为
        'ситуация',  # 局势
        'продолжать',  # 继续
        'готовить',  # 准备
        'агентство',  # 通讯社
        'заявление', 'заявить',  # 声明
        'написать',  # 写道
    }

    # 追加顽固代词和系动词（lemma 形式）
    pronoun_verb_stopwords = {
        # 代词
        'свой', 'который', 'весь', 'этот', 'такой', 'наш', 'ваш',
        'их', 'его', 'её', 'тот', 'один', 'два', 'три',
        # 系动词和情态动词
        'быть', 'являться', 'мочь',
        # 介词和连词
        'против', 'также', 'между',
        # 时间词
        'время', 'год', 'день',
        # 其他
        'это', 'если', 'любой', 'постоянно',
    }

    stopwords_set.update(news_stopwords)
    stopwords_set.update(pronoun_verb_stopwords)

    print(f"✓ 追加 {len(news_stopwords)} 个新闻报道类停用词")
    print(f"✓ 追加 {len(pronoun_verb_stopwords)} 个代词/系动词停用词")
    print(f"✓ 俄语停用词总数：{len(stopwords_set)}")

    return stopwords_set


# 加载停用词（在模块导入时执行）
print("\n" + "="*60)
print("初始化停用词表...")
print("="*60)
CHN_STOPWORDS = load_chinese_stopwords()
print()
RUS_STOPWORDS = load_russian_stopwords()
print("="*60 + "\n")

# 保留的俄语词性（只保留实词）
RUS_KEEP_POS = {'NOUN', 'VERB', 'ADJF', 'ADJS', 'ADVB', 'INFN', 'PRTF', 'PRTS'}

# 字体配置（防止中文乱码）
FONT_PATH = 'simhei.ttf'  # 黑体，Windows系统自带
# 如果找不到字体，尝试以下路径：
# Windows: C:/Windows/Fonts/simhei.ttf
# Mac: /System/Library/Fonts/PingFang.ttc
# Linux: /usr/share/fonts/truetype/wqy/wqy-microhei.ttc

# 输出目录
OUTPUT_DIR = 'outputs_3'

# ============================================================================
# 工具函数
# ============================================================================

def is_chinese(text: str) -> bool:
    """判断文本是否主要为中文"""
    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
    return chinese_chars > len(text) * 0.3


def is_russian(text: str) -> bool:
    """判断文本是否主要为俄文"""
    russian_chars = len(re.findall(r'[а-яА-ЯёЁ]', text))
    return russian_chars > len(text) * 0.3


def extract_dates(text: str) -> List[str]:
    """从文本中提取日期（用于事件划分）"""
    # 匹配常见日期格式：1931年9月18日、9月18日、1931-09-18、7.9、7.10等
    date_patterns = [
        r'\d{4}年\d{1,2}月\d{1,2}日',
        r'\d{1,2}月\d{1,2}日',
        r'\d{4}-\d{2}-\d{2}',
        r'\d{2}\.\d{2}\.\d{4}',
        r'\b\d{1,2}\.\d{1,2}\b',  # 新增：支持 7.9、7.10 格式
    ]
    dates = []
    for pattern in date_patterns:
        dates.extend(re.findall(pattern, text))
    return dates


# ============================================================================
# 文档读取与文本分离
# ============================================================================

def read_docx(file_path: str) -> Tuple[List[str], List[str], List[Dict]]:
    """
    读取Word文档，分离中俄文文本

    返回:
        chinese_texts: 中文段落列表
        russian_texts: 俄文段落列表
        events: 按日期/段落划分的事件列表
    """
    print(f"\n{'='*60}")
    print(f"正在读取文档: {file_path}")
    print(f"{'='*60}")

    doc = Document(file_path)
    chinese_texts = []
    russian_texts = []
    events = []

    current_event = {'date': None, 'chinese': [], 'russian': []}

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # 清洗：剔除类似 <ТАСС> 的前缀标签
        text = re.sub(r'<[^>]+>', '', text).strip()
        if not text:
            continue

        # 尝试提取日期（作为事件分隔符）
        dates = extract_dates(text)
        if dates:
            # 保存上一个事件
            if current_event['chinese'] or current_event['russian']:
                events.append(current_event)
            # 开始新事件
            current_event = {
                'date': dates[0],
                'chinese': [],
                'russian': []
            }

        # 判断语言并分类
        if is_chinese(text):
            chinese_texts.append(text)
            current_event['chinese'].append(text)
        elif is_russian(text):
            russian_texts.append(text)
            current_event['russian'].append(text)

    # 保存最后一个事件
    if current_event['chinese'] or current_event['russian']:
        events.append(current_event)

    print(f"✓ 读取完成:")
    print(f"  - 中文段落数: {len(chinese_texts)}")
    print(f"  - 俄文段落数: {len(russian_texts)}")
    print(f"  - 事件数量: {len(events)}")

    return chinese_texts, russian_texts, events


# ============================================================================
# 中文文本处理
# ============================================================================

def process_chinese_text(texts: List[str]) -> List[Tuple[str, str]]:
    """
    处理中文文本：分词、停用词过滤、同义词合并

    参数:
        texts: 中文文本列表

    返回:
        处理后的 [(词汇, 词性), ...] 列表
    """
    print(f"\n{'='*60}")
    print("开始处理中文文本...")
    print(f"{'='*60}")

    all_words = []

    for text in texts:
        # 清洗：去除换行符、多余空格
        text = re.sub(r'\s+', '', text)

        # 分词并标注词性
        words = pseg.cut(text)

        for word, pos in words:
            word = word.strip()

            # 过滤：长度、停用词、纯数字、非中文
            if len(word) < 2:
                continue
            if word in CHN_STOPWORDS:
                continue
            if re.match(r'^\d+$', word):  # 严格过滤纯数字
                continue
            if not re.search(r'[\u4e00-\u9fff]', word):  # 必须包含汉字
                continue

            # 同义词合并
            word = CHN_SYNONYMS.get(word, word)
            all_words.append((word, pos))

    print(f"✓ 中文处理完成，提取有效词汇: {len(all_words)} 个")
    return all_words


# ============================================================================
# 俄语文本处理
# ============================================================================

def process_russian_text(texts: List[str]) -> List[Tuple[str, str]]:
    """
    处理俄语文本：词形还原、词性过滤、停用词过滤、同义词合并

    参数:
        texts: 俄文文本列表

    返回:
        处理后的 [(词汇, 词性), ...] 列表（已还原为基本形式）
    """
    print(f"\n{'='*60}")
    print("开始处理俄语文本...")
    print(f"{'='*60}")

    morph = pymorphy3.MorphAnalyzer()
    all_words = []

    for text in texts:
        # 分词（按空格和标点）
        words = re.findall(r'[а-яА-ЯёЁ]+', text.lower())

        for word in words:
            # 过滤长度
            if len(word) < 3:
                continue

            # 词形还原
            parsed = morph.parse(word)[0]
            lemma = parsed.normal_form  # 基本形式
            pos = parsed.tag.POS  # 词性

            # 词性过滤（只保留实词）
            if pos not in RUS_KEEP_POS:
                continue

            # 停用词过滤
            if lemma in RUS_STOPWORDS:
                continue

            # 同义词合并
            lemma = RUS_SYNONYMS.get(lemma, lemma)
            all_words.append((lemma, pos))

    print(f"✓ 俄语处理完成，提取有效词汇: {len(all_words)} 个")
    return all_words


# ============================================================================
# 频次统计
# ============================================================================

def calculate_frequencies(words: List[Tuple[str, str]], top_n: int = 20) -> pd.DataFrame:
    """
    计算词频统计

    参数:
        words: [(词汇, 词性), ...] 列表
        top_n: 返回前N个高频词

    返回:
        包含排名、词汇、词性、绝对频次的DataFrame
    """
    # 统计词频（只统计词汇，不考虑词性）
    word_list = [w[0] for w in words]
    counter = Counter(word_list)
    most_common = counter.most_common(top_n)

    # 为每个高频词找到最常见的词性
    word_pos_map = {}
    for word, pos in words:
        if word in [w for w, _ in most_common]:
            if word not in word_pos_map:
                word_pos_map[word] = []
            word_pos_map[word].append(pos)

    # 选择每个词最常见的词性
    word_main_pos = {}
    for word, pos_list in word_pos_map.items():
        word_main_pos[word] = Counter(pos_list).most_common(1)[0][0]

    df = pd.DataFrame(most_common, columns=['词汇', '绝对频次'])
    df.insert(0, '排名', range(1, len(df) + 1))
    df.insert(2, '词性', df['词汇'].map(word_main_pos))

    return df


def calculate_normalized_frequencies(
    events: List[Dict],
    language: str,
    top_words: List[str]
) -> Dict[str, float]:
    """
    计算标准化频次（每万词频次）

    参数:
        events: 事件列表
        language: 'chinese' 或 'russian'
        top_words: 需要计算标准化频次的词汇列表

    返回:
        {词汇: 标准化频次} 字典
    """
    word_event_counts = defaultdict(int)  # 词汇在各事件中的出现次数
    event_total_words = []  # 各事件的总词数

    for event in events:
        texts = event.get('chinese' if language == 'chinese' else 'russian', [])
        if not texts:
            continue

        # 处理该事件的文本
        if language == 'chinese':
            word_tuples = process_chinese_text(texts)
        else:
            word_tuples = process_russian_text(texts)

        # 提取词汇列表
        words = [w[0] for w in word_tuples]
        event_total = len(words)
        if event_total == 0:
            continue

        event_total_words.append(event_total)

        # 统计top词在该事件中的出现次数
        word_counts = Counter(words)
        for word in top_words:
            if word in word_counts:
                # 标准化频次 = (词频 / 事件总词数) * 10000
                normalized = (word_counts[word] / event_total) * 10000
                word_event_counts[word] += normalized

    # 计算平均标准化频次
    num_events = len([e for e in events if e.get('chinese' if language == 'chinese' else 'russian')])
    if num_events == 0:
        return {word: 0 for word in top_words}

    return {word: word_event_counts[word] / num_events for word in top_words}


# ============================================================================
# 可视化
# ============================================================================

def generate_wordcloud(words: List[Tuple[str, str]], title: str, output_file: str, output_dir: str = None, font_path: str = None):
    """
    生成词云图

    参数:
        words: [(词汇, 词性), ...] 列表
        title: 图表标题
        output_file: 输出文件名
        output_dir: 输出目录，默认使用全局 OUTPUT_DIR
        font_path: 字体路径，默认使用全局 FONT_PATH
    """
    print(f"\n生成词云图: {title}")

    # 词频统计（只统计词汇）
    word_list = [w[0] for w in words]
    word_freq = Counter(word_list)

    # 生成词云
    wordcloud = WordCloud(
        font_path=font_path or FONT_PATH,
        width=1200,
        height=800,
        background_color='white',
        max_words=100,
        relative_scaling=0.5,
        colormap='viridis'
    ).generate_from_frequencies(word_freq)

    # 绘图
    plt.figure(figsize=(15, 10))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis('off')
    # plt.title(title, fontproperties='SimHei', fontsize=20, pad=20) # 不要标题
    plt.tight_layout(pad=0)

    # 保存
    save_dir = output_dir if output_dir else OUTPUT_DIR
    output_path = os.path.join(save_dir, output_file)
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()

    print(f"✓ 词云图已保存: {output_path}")


def generate_bar_chart(df: pd.DataFrame, title: str, output_file: str, output_dir: str = None, use_times_for_xticks: bool = False):
    """
    生成高频词柱状图

    参数:
        df: 包含词汇和标准化频次的DataFrame
        title: 图表标题
        output_file: 输出文件名
        output_dir: 输出目录，默认使用全局 OUTPUT_DIR
        use_times_for_xticks: 是否对X轴刻度使用Times New Roman字体（用于俄文）
    """
    print(f"\n生成柱状图: {title}")

    # 设置中文字体（用于标题和坐标轴标签）
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False

    # 绘图
    fig, ax = plt.subplots(figsize=(14, 8))

    bars = ax.bar(
        range(len(df)),
        df['标准化频次'],
        color='steelblue',
        alpha=0.8,
        edgecolor='black',
        linewidth=0.5
    )

    # 设置x轴标签
    ax.set_xticks(range(len(df)))

    if use_times_for_xticks:
        # 俄文：X轴刻度使用 Times New Roman
        ax.set_xticklabels(
            df['词汇'],
            rotation=45,
            ha='right',
            fontsize=11,
            fontfamily='Times New Roman'
        )
    else:
        # 中文：使用默认中文字体
        ax.set_xticklabels(df['词汇'], rotation=45, ha='right', fontsize=11)

    # 设置标题和标签（中文，使用SimHei）
    ax.set_title(title, fontsize=16, fontweight='bold', pad=20)
    ax.set_xlabel('词汇', fontsize=13, fontweight='bold')
    ax.set_ylabel('标准化频次（每万词）', fontsize=13, fontweight='bold')

    # 添加网格
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    ax.set_axisbelow(True)

    # 在柱子上方显示数值
    for i, bar in enumerate(bars):
        height = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            height,
            f'{height:.1f}',
            ha='center',
            va='bottom',
            fontsize=9
        )

    plt.tight_layout()

    # 保存
    save_dir = output_dir if output_dir else OUTPUT_DIR
    output_path = os.path.join(save_dir, output_file)
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()

    print(f"✓ 柱状图已保存: {output_path}")


# ============================================================================
# 主流程
# ============================================================================

# ============================================================================
# 工具函数（修改部分）
# ============================================================================

def ensure_output_dir(output_dir: str) -> None:
    """确保输出目录存在，不存在则创建（支持多级路径）"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"  [目录已创建] {output_dir}")


def get_doc_output_dir(base_name: str) -> str:
    """
    根据文档名称，返回其专属输出子目录路径。
    例如：'七七事变资料' -> 'outputs/七七事变资料'
    """
    return os.path.join(OUTPUT_DIR, base_name)


# ============================================================================
# 主流程（重构部分）
# ============================================================================

def main():
    """批量处理 Data 文件夹下所有 docx 文件的主流程"""
    print("=" * 60)
    print("  历史语料 NLP 批量分析脚本")
    print("=" * 60)

    # ── 第一点：扫描 Data 文件夹，收集所有 docx 文件 ──────────────────────
    data_dir = 'Data'
    if not os.path.isdir(data_dir):
        print(f"[错误] 未找到数据目录：{data_dir}")
        return

    docx_files = sorted([
        os.path.join(data_dir, f)
        for f in os.listdir(data_dir)
        if f.lower().endswith('.docx')
    ])

    if not docx_files:
        print(f"[警告] {data_dir} 目录下未找到任何 .docx 文件，程序退出。")
        return

    total_files = len(docx_files)
    print(f"\n共发现 {total_files} 个 docx 文件，开始批量处理...\n")

    # 用于最终汇总报告
    success_list = []
    failed_list  = []

    # ── 遍历每个文档 ────────────────────────────────────────────────────────
    for idx, docx_path in enumerate(docx_files, start=1):
        base_name  = os.path.splitext(os.path.basename(docx_path))[0]

        # ── 第二点：为每个文档创建同名独立子目录 ──────────────────────────
        output_dir = get_doc_output_dir(base_name)

        print(f"\n{'─'*60}")
        print(f"[{idx}/{total_files}] 正在处理：{base_name}.docx")
        print(f"            输出目录：{output_dir}")
        print(f"{'─'*60}")

        # 创建该文档专属的输出目录
        ensure_output_dir(output_dir)

        try:
            # 1. 读取文档
            chinese_texts, russian_texts, events = read_docx(docx_path)

            # 2. 中文词频统计
            chinese_words    = process_chinese_text(chinese_texts)
            chinese_freq_df  = calculate_frequencies(chinese_words, top_n=20)

            # 3. 俄文词频统计
            russian_words    = process_russian_text(russian_texts)
            russian_freq_df  = calculate_frequencies(russian_words, top_n=20)

            # 4. 标准化频次
            print(f"\n{'='*60}")
            print("计算标准化频次...")
            print(f"{'='*60}")

            chinese_normalized = calculate_normalized_frequencies(
                events, 'chinese', chinese_freq_df['词汇'].tolist()
            )
            russian_normalized = calculate_normalized_frequencies(
                events, 'russian', russian_freq_df['词汇'].tolist()
            )

            chinese_freq_df['标准化频次'] = chinese_freq_df['词汇'].map(chinese_normalized)
            russian_freq_df['标准化频次'] = russian_freq_df['词汇'].map(russian_normalized)

            chinese_freq_df.insert(1, '语言', '中文')
            russian_freq_df.insert(1, '语言', '俄文')

            print("标准化频次计算完成")

            # ── 第三点：词云、柱状图、Excel 均写入该文档的独立子目录 ────────

            # 5. 生成词云（→ output_dir/wordcloud_chinese.png 等）
            print(f"\n{'='*60}")
            print("生成词云图...")
            print(f"{'='*60}")

            RUS_FONT = 'C:/Windows/Fonts/times.ttf'

            generate_wordcloud(
                chinese_words, '中文词汇词云图',
                'wordcloud_chinese.png', output_dir=output_dir
            )
            generate_wordcloud(
                russian_words, '俄文词汇词云图',
                'wordcloud_russian.png', output_dir=output_dir,
                font_path=RUS_FONT
            )

            # 6. 生成柱状图（→ output_dir/barchart_chinese.png 等）
            chinese_bar_df = chinese_freq_df.sort_values(
                '标准化频次', ascending=False
            ).reset_index(drop=True)
            russian_bar_df = russian_freq_df.sort_values(
                '标准化频次', ascending=False
            ).reset_index(drop=True)

            generate_bar_chart(
                chinese_bar_df, '中文 Top20 标准化频次',
                'barchart_chinese.png', output_dir=output_dir,
                use_times_for_xticks=False
            )
            generate_bar_chart(
                russian_bar_df, '俄文 Top20 标准化频次',
                'barchart_russian.png', output_dir=output_dir,
                use_times_for_xticks=True
            )

            # 7. 导出 Excel（→ output_dir/result.xlsx）
            print(f"\n{'='*60}")
            print("导出结果表格...")
            print(f"{'='*60}")

            print("\n翻译俄文词汇...")
            translator = GoogleTranslator(source='ru', target='zh-CN')

            def translate_words_robust(words):
                """优先使用本地词典，不足时调用 API"""
                results         = []
                api_translations = []
                for w in words:
                    if w in RU_TO_ZH_TRANS:
                        results.append(RU_TO_ZH_TRANS[w])
                    else:
                        try:
                            trans = translator.translate(w)
                            results.append(trans)
                            api_translations.append((w, trans))
                        except Exception as e:
                            print(f"  翻译失败：{w} - {e}")
                            results.append('')

                if api_translations:
                    print("\n" + "=" * 60)
                    print("以下词汇通过 API 翻译，建议补充到 RU_TO_ZH_TRANS 字典：")
                    print("=" * 60)
                    for ru, zh in api_translations:
                        print(f"    '{ru}': '{zh}',")
                    print("=" * 60 + "\n")

                return results

            russian_freq_df['中文翻译'] = translate_words_robust(
                russian_freq_df['词汇'].tolist()
            )
            chinese_freq_df['中文翻译'] = chinese_freq_df['词汇']

            def build_export_df(df, lang):
                return pd.DataFrame({
                    '语言':     lang,
                    '词汇':     df['词汇'].values,
                    '词性':     df['词性'].values if lang == '中文' else '',
                    '中文翻译': df['中文翻译'].values,
                    '频次':     df['绝对频次'].values,
                    '排名':     df['排名'].values,
                    '标准化频次': df['标准化频次'].values,
                    '备注':     '',
                })

            combined_df = pd.concat(
                [build_export_df(chinese_freq_df, '中文'),
                 build_export_df(russian_freq_df, '俄文')],
                ignore_index=True
            )

            # ── 第三点（关键）：Excel 路径使用当前文档专属 output_dir ──────
            excel_path = os.path.join(output_dir, 'result.xlsx')
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                combined_df.to_excel(
                    writer, sheet_name='综合词频表', index=False
                )
                chinese_freq_df.to_excel(
                    writer, sheet_name='中文词频', index=False
                )
                russian_freq_df.to_excel(
                    writer, sheet_name='俄文词频', index=False
                )

                events_data = [
                    {
                        '事件序号': i,
                        '日期':     ev['date'] or '未知',
                        '中文段落数': len(ev['chinese']),
                        '俄文段落数': len(ev['russian'])
                    }
                    for i, ev in enumerate(events, 1)
                ]
                pd.DataFrame(events_data).to_excel(
                    writer, sheet_name='事件列表', index=False
                )

            print(f"Excel 已保存：{excel_path}")

            # 8. 单文件摘要
            print(f"\n{'='*60}")
            print(f"「{base_name}」处理完成")
            print(f"{'='*60}")
            print(f"\n中文 Top5 词汇")
            for _, row in chinese_freq_df.head(5).iterrows():
                print(f"  {row['排名']}. {row['词汇']:<10} "
                      f"频次：{row['绝对频次']:<6} "
                      f"中文翻译：{row['中文翻译']:>4}  "
                      f"标准化频次：{row['标准化频次']:>6.2f}")

            print(f"\n俄文 Top5 词汇")
            for _, row in russian_freq_df.head(5).iterrows():
                print(f"  {row['排名']}. {row['词汇']:<15} "
                      f"频次：{row['绝对频次']:<6} "
                      f"中文翻译：{row['中文翻译']:>4}  "
                      f"标准化频次：{row['标准化频次']:>6.2f}")

            print(f"\n所有输出文件位于：{output_dir}")

            success_list.append(base_name)

        except Exception as e:
            print(f"\n[错误] 处理「{base_name}」时发生异常：{e}")
            failed_list.append((base_name, str(e)))

    # ── 第四点：全局运行报告 ───────────────────────────────────────────────
    print(f"\n{'='*60}")
    print("  批量处理完毕 —— 总体运行报告")
    print(f"{'='*60}")
    print(f"  总文件数  ：{total_files}")
    print(f"  成功处理  ：{len(success_list)} 个")
    print(f"  失败文件  ：{len(failed_list)} 个")

    if success_list:
        print(f"\n  ✓ 成功列表：")
        for name in success_list:
            print(f"      {name}  →  {get_doc_output_dir(name)}")

    if failed_list:
        print(f"\n  ✗ 失败列表（请检查后重试）：")
        for name, reason in failed_list:
            print(f"      {name}：{reason}")

    print(f"\n  根输出目录：{OUTPUT_DIR}")
    print(f"{'='*60}\n")


if __name__ == '__main__':
    main()