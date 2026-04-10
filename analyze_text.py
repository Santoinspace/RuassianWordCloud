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

# 中文停用词表（基础版）
CHN_STOPWORDS = {
    '的', '了', '在', '是', '我', '有', '和', '就', '不', '人', '都', '一', '一个',
    '上', '也', '很', '到', '说', '要', '去', '你', '会', '着', '没有', '看', '好',
    '自己', '这', '那', '里', '为', '与', '及', '等', '之', '于', '对', '从', '以',
    '但', '而', '或', '因', '所', '由', '其', '被', '将', '把', '向', '给', '让',
    '年', '月', '日', '时', '分', '点', '号', '第', '个', '些', '此', '该', '各', '塔斯社',
    # 代词和量词
    '一些', '这种', '那种', '这样', '那样', '什么', '怎么', '如何', '哪些', '多少',
    # 领域泛义词（新闻报道类低价值词）
    '发生', '问题', '进行', '要求', '方面', '行动', '认为', '局势', '继续', '准备',
    '情况', '表示', '指出', '提出', '作出', '采取', '开始', '结束', '出现', '成为',
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
OUTPUT_DIR = 'outputs\\xian_analysis'

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

def generate_wordcloud(words: List[Tuple[str, str]], title: str, output_file: str, font_path: str = None):
    """
    生成词云图

    参数:
        words: [(词汇, 词性), ...] 列表
        title: 图表标题
        output_file: 输出文件名
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
    output_path = os.path.join(OUTPUT_DIR, output_file)
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()

    print(f"✓ 词云图已保存: {output_path}")


def generate_bar_chart(df: pd.DataFrame, title: str, output_file: str, use_times_for_xticks: bool = False):
    """
    生成高频词柱状图

    参数:
        df: 包含词汇和标准化频次的DataFrame
        title: 图表标题
        output_file: 输出文件名
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
    print("抗日事变历史语料NLP分析系统")
    print("="*60)

    # 确保输出目录存在
    ensure_output_dir()

    # 1. 读取文档
    docx_path = 'Data/西安事变资料.docx'
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

    # 俄语词云明确使用 Times New Roman 字体
    RUS_FONT = 'C:/Windows/Fonts/times.ttf'

    generate_wordcloud(chinese_words, '中文高频词词云图', 'wordcloud_chinese.png')
    generate_wordcloud(russian_words, '俄文高频词词云图', 'wordcloud_russian.png', font_path=RUS_FONT)

    # 6. 生成柱状图
    generate_bar_chart(chinese_freq_df, '中文Top20高频词分布', 'barchart_chinese.png', use_times_for_xticks=False)
    generate_bar_chart(russian_freq_df, '俄文Top20高频词分布', 'barchart_russian.png', use_times_for_xticks=True)

    # 7. 导出Excel
    print(f"\n{'='*60}")
    print("导出统计数据...")
    print(f"{'='*60}")

    # 为俄文词汇自动翻译中文（优先使用字典，其次API）
    print("\n正在翻译俄文词汇...")
    translator = GoogleTranslator(source='ru', target='zh-CN')

    def translate_words_robust(words):
        """鲁棒的翻译函数：优先字典，其次API"""
        results = []
        api_translations = []  # 记录通过API获取的翻译

        for w in words:
            # 1. 先查字典
            if w in RU_TO_ZH_TRANS:
                results.append(RU_TO_ZH_TRANS[w])
            else:
                # 2. 字典中没有，调用API
                try:
                    trans = translator.translate(w)
                    results.append(trans)
                    api_translations.append((w, trans))
                except Exception as e:
                    print(f"  ⚠ 翻译失败: {w} - {e}")
                    results.append('')

        # 3. 打印API翻译结果，提醒用户添加到字典
        if api_translations:
            print("\n" + "="*60)
            print("⚠ 以下翻译通过API获取，建议添加到 RU_TO_ZH_TRANS 字典：")
            print("="*60)
            for ru, zh in api_translations:
                print(f"    '{ru}': '{zh}',")
            print("="*60 + "\n")

        return results

    russian_freq_df['中文翻译'] = translate_words_robust(russian_freq_df['词汇'].tolist())
    chinese_freq_df['中文翻译'] = chinese_freq_df['词汇']  # 中文词汇本身即为翻译

    # 构建统一格式：语言 | 排名 | 俄文 | 中文翻译 | 频次 | 词性 | 标准化频次 | 语义分类
    def build_export_df(df, lang):
        return pd.DataFrame({
            '语言':     lang,
            '排名':     df['排名'].values,
            '俄文':     df['词汇'].values if lang == '俄文' else '',
            '中文翻译': df['中文翻译'].values,
            '频次':     df['绝对频次'].values,
            '词性':     df['词性'].values,
            '标准化频次': df['标准化频次'].values,
            '语义分类': '',
        })

    combined_df = pd.concat(
        [build_export_df(chinese_freq_df, '中文'), build_export_df(russian_freq_df, '俄文')],
        ignore_index=True
    )

    # 保存Excel
    excel_path = os.path.join(OUTPUT_DIR, 'result.xlsx')
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        combined_df.to_excel(writer, sheet_name='高频词统计', index=False)
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
              f"词性: {row['词性']:<6} "
              f"绝对频次: {row['绝对频次']:>4}  "
              f"标准化频次: {row['标准化频次']:>6.2f}")

    print(f"\n【俄文Top5高频词】")
    for _, row in russian_freq_df.head(5).iterrows():
        print(f"  {row['排名']}. {row['词汇']:<15} "
              f"词性: {row['词性']:<6} "
              f"绝对频次: {row['绝对频次']:>4}  "
              f"标准化频次: {row['标准化频次']:>6.2f}")

    print(f"\n{'='*60}")
    print("所有文件已保存到 output/ 目录")
    print(f"{'='*60}\n")


if __name__ == '__main__':
    main()
