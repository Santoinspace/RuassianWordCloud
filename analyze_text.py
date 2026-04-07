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
import pandas as pd
import matplotlib.pyplot as plt
from wordcloud import WordCloud
from docx import Document
import pymorphy3
from nltk.corpus import stopwords
import nltk

# ============================================================================
# 全局配置区域
# ============================================================================

# 中文同义词映射字典（可手动维护）
CHN_SYNONYMS = {
    '日本军': '日军',
    '日本人': '日本',
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

# 中文停用词表（基础版）
CHN_STOPWORDS = {
    '的', '了', '在', '是', '我', '有', '和', '就', '不', '人', '都', '一', '一个',
    '上', '也', '很', '到', '说', '要', '去', '你', '会', '着', '没有', '看', '好',
    '自己', '这', '那', '里', '为', '与', '及', '等', '之', '于', '对', '从', '以',
    '但', '而', '或', '因', '所', '由', '其', '被', '将', '把', '向', '给', '让',
    '年', '月', '日', '时', '分', '点', '号', '第', '个', '些', '此', '该', '各',
}

# 俄语停用词（需要下载nltk数据）
try:
    nltk.data.find('corpora/stopwords')
except LookupError:
    nltk.download('stopwords')

RUS_STOPWORDS = set(stopwords.words('russian'))

# 俄语额外停用词（补充常见功能词）
RUS_STOPWORDS.update({
    'это', 'быть', 'весь', 'свой', 'мочь', 'который', 'такой', 'наш', 'ваш',
    'их', 'его', 'её', 'этот', 'тот', 'один', 'два', 'три', 'год', 'день',
})

# 保留的俄语词性（只保留实词）
RUS_KEEP_POS = {'NOUN', 'VERB', 'ADJF', 'ADJS', 'ADVB', 'INFN', 'PRTF', 'PRTS'}

# 字体配置（防止中文乱码）
FONT_PATH = 'simhei.ttf'  # 黑体，Windows系统自带
# 如果找不到字体，尝试以下路径：
# Windows: C:/Windows/Fonts/simhei.ttf
# Mac: /System/Library/Fonts/PingFang.ttc
# Linux: /usr/share/fonts/truetype/wqy/wqy-microhei.ttc

# 输出目录
OUTPUT_DIR = 'output'

# ============================================================================
# 工具函数
# ============================================================================

def ensure_output_dir():
    """确保输出目录存在"""
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"✓ 创建输出目录: {OUTPUT_DIR}")


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
    # 匹配常见日期格式：1931年9月18日、9月18日、1931-09-18等
    date_patterns = [
        r'\d{4}年\d{1,2}月\d{1,2}日',
        r'\d{1,2}月\d{1,2}日',
        r'\d{4}-\d{2}-\d{2}',
        r'\d{2}\.\d{2}\.\d{4}',
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

def process_chinese_text(texts: List[str]) -> List[str]:
    """
    处理中文文本：分词、停用词过滤、同义词合并

    参数:
        texts: 中文文本列表

    返回:
        处理后的词汇列表
    """
    print(f"\n{'='*60}")
    print("开始处理中文文本...")
    print(f"{'='*60}")

    all_words = []

    for text in texts:
        # 分词
        words = jieba.lcut(text)

        for word in words:
            # 过滤：长度、停用词、标点
            if len(word) < 2:
                continue
            if word in CHN_STOPWORDS:
                continue
            if re.match(r'^[^\u4e00-\u9fff]+$', word):  # 非中文字符
                continue

            # 同义词合并
            word = CHN_SYNONYMS.get(word, word)
            all_words.append(word)

    print(f"✓ 中文处理完成，提取有效词汇: {len(all_words)} 个")
    return all_words


# ============================================================================
# 俄语文本处理
# ============================================================================

def process_russian_text(texts: List[str]) -> List[str]:
    """
    处理俄语文本：词形还原、词性过滤、停用词过滤、同义词合并

    参数:
        texts: 俄文文本列表

    返回:
        处理后的词汇列表（已还原为基本形式）
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
            all_words.append(lemma)

    print(f"✓ 俄语处理完成，提取有效词汇: {len(all_words)} 个")
    return all_words


# ============================================================================
# 频次统计
# ============================================================================

def calculate_frequencies(words: List[str], top_n: int = 20) -> pd.DataFrame:
    """
    计算词频统计

    参数:
        words: 词汇列表
        top_n: 返回前N个高频词

    返回:
        包含排名、词汇、绝对频次的DataFrame
    """
    counter = Counter(words)
    most_common = counter.most_common(top_n)

    df = pd.DataFrame(most_common, columns=['词汇', '绝对频次'])
    df.insert(0, '排名', range(1, len(df) + 1))

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
            words = process_chinese_text(texts)
        else:
            words = process_russian_text(texts)

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

def generate_wordcloud(words: List[str], title: str, output_file: str):
    """
    生成词云图

    参数:
        words: 词汇列表
        title: 图表标题
        output_file: 输出文件名
    """
    print(f"\n生成词云图: {title}")

    # 词频统计
    word_freq = Counter(words)

    # 生成词云
    wordcloud = WordCloud(
        font_path=FONT_PATH,
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
    plt.title(title, fontproperties='SimHei', fontsize=20, pad=20)
    plt.tight_layout(pad=0)

    # 保存
    output_path = os.path.join(OUTPUT_DIR, output_file)
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()

    print(f"✓ 词云图已保存: {output_path}")


def generate_bar_chart(df: pd.DataFrame, title: str, output_file: str):
    """
    生成高频词柱状图

    参数:
        df: 包含词汇和标准化频次的DataFrame
        title: 图表标题
        output_file: 输出文件名
    """
    print(f"\n生成柱状图: {title}")

    # 设置中文字体
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
    ax.set_xticklabels(df['词汇'], rotation=45, ha='right', fontsize=11)

    # 设置标题和标签
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
    output_path = os.path.join(OUTPUT_DIR, output_file)
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()

    print(f"✓ 柱状图已保存: {output_path}")


# ============================================================================
# 主流程
# ============================================================================

def main():
    """主函数"""
    print("\n" + "="*60)
    print("九一八事变历史语料NLP分析系统")
    print("="*60)

    # 确保输出目录存在
    ensure_output_dir()

    # 1. 读取文档
    docx_path = 'Data/九一八事变资料.docx'
    chinese_texts, russian_texts, events = read_docx(docx_path)

    # 2. 处理中文文本
    chinese_words = process_chinese_text(chinese_texts)
    chinese_freq_df = calculate_frequencies(chinese_words, top_n=20)

    # 3. 处理俄语文本
    russian_words = process_russian_text(russian_texts)
    russian_freq_df = calculate_frequencies(russian_words, top_n=20)

    # 4. 计算标准化频次
    print(f"\n{'='*60}")
    print("计算标准化频次...")
    print(f"{'='*60}")

    chinese_normalized = calculate_normalized_frequencies(
        events, 'chinese', chinese_freq_df['词汇'].tolist()
    )
    russian_normalized = calculate_normalized_frequencies(
        events, 'russian', russian_freq_df['词汇'].tolist()
    )

    # 添加标准化频次列
    chinese_freq_df['标准化频次'] = chinese_freq_df['词汇'].map(chinese_normalized)
    russian_freq_df['标准化频次'] = russian_freq_df['词汇'].map(russian_normalized)

    # 添加语言标识
    chinese_freq_df.insert(1, '语言', '中文')
    russian_freq_df.insert(1, '语言', '俄文')

    print(f"✓ 标准化频次计算完成")

    # 5. 生成词云图
    print(f"\n{'='*60}")
    print("生成可视化图表...")
    print(f"{'='*60}")

    generate_wordcloud(chinese_words, '中文高频词词云图', 'wordcloud_chinese.png')
    generate_wordcloud(russian_words, '俄文高频词词云图', 'wordcloud_russian.png')

    # 6. 生成柱状图
    generate_bar_chart(chinese_freq_df, '中文Top20高频词分布', 'barchart_chinese.png')
    generate_bar_chart(russian_freq_df, '俄文Top20高频词分布', 'barchart_russian.png')

    # 7. 导出Excel
    print(f"\n{'='*60}")
    print("导出统计数据...")
    print(f"{'='*60}")

    # 合并中俄文数据
    combined_df = pd.concat([chinese_freq_df, russian_freq_df], ignore_index=True)

    # 添加词性列（预留，可后续手动填写）
    combined_df.insert(4, '词性', '')
    combined_df.insert(5, '语义分类', '')  # 预留：主体类、行动类等

    # 调整列顺序
    combined_df = combined_df[[
        '语言', '排名', '词汇', '词性', '语义分类', '绝对频次', '标准化频次'
    ]]

    # 保存Excel
    excel_path = os.path.join(OUTPUT_DIR, 'result.xlsx')
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        combined_df.to_excel(writer, sheet_name='高频词统计', index=False)

        # 分别保存中俄文数据
        chinese_freq_df.to_excel(writer, sheet_name='中文统计', index=False)
        russian_freq_df.to_excel(writer, sheet_name='俄文统计', index=False)

        # 保存事件划分信息
        events_data = []
        for i, event in enumerate(events, 1):
            events_data.append({
                '事件编号': i,
                '日期': event['date'] or '未标注',
                '中文段落数': len(event['chinese']),
                '俄文段落数': len(event['russian'])
            })
        events_df = pd.DataFrame(events_data)
        events_df.to_excel(writer, sheet_name='事件划分', index=False)

    print(f"✓ Excel文件已保存: {excel_path}")

    # 8. 输出统计摘要
    print(f"\n{'='*60}")
    print("分析完成！统计摘要：")
    print(f"{'='*60}")
    print(f"\n【中文Top5高频词】")
    for _, row in chinese_freq_df.head(5).iterrows():
        print(f"  {row['排名']}. {row['词汇']:<10} "
              f"绝对频次: {row['绝对频次']:>4}  "
              f"标准化频次: {row['标准化频次']:>6.2f}")

    print(f"\n【俄文Top5高频词】")
    for _, row in russian_freq_df.head(5).iterrows():
        print(f"  {row['排名']}. {row['词汇']:<15} "
              f"绝对频次: {row['绝对频次']:>4}  "
              f"标准化频次: {row['标准化频次']:>6.2f}")

    print(f"\n{'='*60}")
    print("所有文件已保存到 output/ 目录")
    print(f"{'='*60}\n")


if __name__ == '__main__':
    main()
