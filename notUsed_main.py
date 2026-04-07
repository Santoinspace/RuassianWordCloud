"""
中俄双语历史语料 NLP 分析与可视化工具
功能：词频统计、词云生成、柱状图绘制、Excel导出
"""

import re
from collections import Counter
from pathlib import Path
import docx
import jieba
import pymorphy3
import pandas as pd
import matplotlib.pyplot as plt
from wordcloud import WordCloud
import matplotlib.font_manager as fm

# ==================== 全局配置 ====================

# 同义词映射字典（可手动维护）
RUS_SYNONYMS = {
    'японцы': 'японский',
    'японец': 'японский',
    'япония': 'японский',
    'китайцы': 'китайский',
    'китаец': 'китайский',
    # 可继续添加...
}

CHN_SYNONYMS = {
    '日本军': '日军',
    '日本人': '日本',
    '中国人': '中国',
    # 可继续添加...
}

# 俄语停用词表（常见语法功能词）
RUS_STOPWORDS = {
    'и', 'в', 'на', 'с', 'по', 'к', 'из', 'о', 'от', 'для', 'у', 'за', 'до', 'при',
    'это', 'как', 'что', 'который', 'весь', 'этот', 'тот', 'наш', 'свой', 'мой', 'твой',
    'он', 'она', 'оно', 'они', 'я', 'ты', 'мы', 'вы', 'его', 'её', 'их', 'себя',
    'быть', 'есть', 'был', 'была', 'было', 'были', 'будет', 'будут',
    'не', 'ни', 'да', 'нет', 'же', 'ли', 'бы', 'только', 'уже', 'еще', 'ещё',
    'а', 'но', 'или', 'если', 'то', 'так', 'там', 'тут', 'где', 'когда', 'почему',
    'год', 'года', 'лет',  # 时间词
}

# 中文停用词表（基础版）
CHN_STOPWORDS = {
    '的', '了', '在', '是', '我', '有', '和', '就', '不', '人', '都', '一', '一个',
    '上', '也', '很', '到', '说', '要', '去', '你', '会', '着', '没有', '看', '好',
    '自己', '这', '那', '里', '就是', '为', '与', '及', '等', '但', '而', '或',
    '年', '月', '日', '时', '分', '个', '中', '之', '于', '对', '从', '以',
}

# 保留的词性（俄语）
RUS_VALID_POS = {'NOUN', 'VERB', 'ADJF', 'ADJS', 'ADVB', 'INFN', 'PRTF', 'PRTS'}

# 输出目录
OUTPUT_DIR = Path('output')
OUTPUT_DIR.mkdir(exist_ok=True)


# ==================== 工具函数 ====================

def get_system_font():
    """自动获取系统中支持中俄文的字体"""
    # Windows 常见字体
    fonts = ['SimHei', 'Microsoft YaHei', 'SimSun', 'Arial Unicode MS']
    for font in fonts:
        try:
            font_path = fm.findfont(fm.FontProperties(family=font))
            if font_path:
                return font_path
        except:
            continue
    # 如果都找不到，返回默认
    return fm.findfont(fm.FontProperties())


def read_docx(file_path):
    """读取Word文档，返回所有段落文本"""
    doc = docx.Document(file_path)
    paragraphs = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    return paragraphs


def is_russian(text):
    """判断文本是否为俄语（包含西里尔字母）"""
    return bool(re.search(r'[а-яА-ЯёЁ]', text))


def is_chinese(text):
    """判断文本是否为中文"""
    return bool(re.search(r'[\u4e00-\u9fff]', text))


def separate_languages(paragraphs):
    """
    分离中俄文文本
    返回：(中文段落列表, 俄文段落列表, 事件分段信息)
    """
    chinese_texts = []
    russian_texts = []
    events = []  # 用于存储按日期/段落划分的事件

    current_event = {'chinese': [], 'russian': [], 'date': None}

    for para in paragraphs:
        # 检测日期标记（如：1931年9月18日）
        date_match = re.search(r'\d{4}年\d{1,2}月\d{1,2}日', para)
        if date_match:
            # 保存上一个事件
            if current_event['chinese'] or current_event['russian']:
                events.append(current_event)
            # 开始新事件
            current_event = {'chinese': [], 'russian': [], 'date': date_match.group()}

        # 分类文本
        if is_russian(para):
            russian_texts.append(para)
            current_event['russian'].append(para)
        elif is_chinese(para):
            chinese_texts.append(para)
            current_event['chinese'].append(para)

    # 保存最后一个事件
    if current_event['chinese'] or current_event['russian']:
        events.append(current_event)

    return chinese_texts, russian_texts, events


# ==================== 中文处理 ====================

def process_chinese(texts, synonyms=None):
    """
    中文文本处理：分词、停用词过滤、同义词合并
    返回：词频Counter对象
    """
    if synonyms is None:
        synonyms = CHN_SYNONYMS

    all_words = []

    for text in texts:
        # 分词
        words = jieba.cut(text)

        for word in words:
            word = word.strip()
            # 过滤：长度、停用词、标点
            if len(word) < 2:
                continue
            if word in CHN_STOPWORDS:
                continue
            if not re.search(r'[\u4e00-\u9fff]', word):  # 必须包含汉字
                continue

            # 同义词替换
            word = synonyms.get(word, word)
            all_words.append(word)

    return Counter(all_words)


# ==================== 俄语处理 ====================

def process_russian(texts, synonyms=None):
    """
    俄语文本处理：词形还原、词性过滤、停用词过滤、同义词合并
    返回：词频Counter对象
    """
    if synonyms is None:
        synonyms = RUS_SYNONYMS

    morph = pymorphy3.MorphAnalyzer()
    all_words = []

    for text in texts:
        # 分词（按空格和标点）
        words = re.findall(r'[а-яА-ЯёЁ]+', text.lower())

        for word in words:
            if len(word) < 3:  # 过滤过短的词
                continue
            if word in RUS_STOPWORDS:
                continue

            # 词形还原
            parsed = morph.parse(word)[0]
            lemma = parsed.normal_form
            pos = parsed.tag.POS

            # 词性过滤（只保留实词）
            if pos not in RUS_VALID_POS:
                continue

            # 同义词替换
            lemma = synonyms.get(lemma, lemma)
            all_words.append(lemma)

    return Counter(all_words)


# ==================== 频次计算 ====================

def calculate_normalized_freq(word_counter, total_words, scale=10000):
    """
    计算标准化频次
    公式：(词频 / 总词数) * scale
    """
    return {word: (count / total_words) * scale
            for word, count in word_counter.items()}


def get_top_words(word_counter, top_n=20):
    """获取Top N高频词"""
    return word_counter.most_common(top_n)


# ==================== 可视化 ====================

def generate_wordcloud(word_freq, output_path, title, font_path=None):
    """生成词云图"""
    if font_path is None:
        font_path = get_system_font()

    wc = WordCloud(
        width=1200,
        height=800,
        background_color='white',
        font_path=font_path,
        max_words=100,
        relative_scaling=0.5,
        colormap='viridis'
    ).generate_from_frequencies(word_freq)

    plt.figure(figsize=(15, 10))
    plt.imshow(wc, interpolation='bilinear')
    plt.axis('off')
    plt.title(title, fontproperties=fm.FontProperties(fname=font_path, size=20))
    plt.tight_layout(pad=0)
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"✓ 词云图已保存: {output_path}")


def generate_bar_chart(top_words, output_path, title, xlabel, ylabel, font_path=None):
    """生成柱状图（使用标准化频次）"""
    if font_path is None:
        font_path = get_system_font()

    font_prop = fm.FontProperties(fname=font_path, size=12)

    words = [w[0] for w in top_words]
    freqs = [w[1] for w in top_words]

    plt.figure(figsize=(14, 8))
    bars = plt.bar(range(len(words)), freqs, color='steelblue', alpha=0.8)

    plt.xlabel(xlabel, fontproperties=font_prop, fontsize=14)
    plt.ylabel(ylabel, fontproperties=font_prop, fontsize=14)
    plt.title(title, fontproperties=font_prop, fontsize=16)
    plt.xticks(range(len(words)), words, rotation=45, ha='right',
               fontproperties=font_prop)
    plt.yticks(fontproperties=font_prop)

    # 在柱子上标注数值
    for bar in bars:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2., height,
                f'{height:.1f}',
                ha='center', va='bottom', fontproperties=font_prop, fontsize=9)

    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"✓ 柱状图已保存: {output_path}")


# ==================== 数据导出 ====================

def export_to_excel(chinese_data, russian_data, output_path):
    """
    导出统计结果到Excel
    格式：语言、排名、词条、绝对频次、标准化频次
    """
    rows = []

    # 中文数据
    for rank, (word, abs_freq, norm_freq) in enumerate(chinese_data, 1):
        rows.append({
            '语言': '中文',
            '排名': rank,
            '词条': word,
            '绝对频次': abs_freq,
            '标准化频次': round(norm_freq, 2),
            '词性': '',  # 预留列，供人工标注
            '分类': ''   # 预留列，供人工标注（如：主体类、行动类等）
        })

    # 俄文数据
    for rank, (word, abs_freq, norm_freq) in enumerate(russian_data, 1):
        rows.append({
            '语言': '俄文',
            '排名': rank,
            '词条': word,
            '绝对频次': abs_freq,
            '标准化频次': round(norm_freq, 2),
            '词性': '',
            '分类': ''
        })

    df = pd.DataFrame(rows)
    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"✓ Excel文件已导出: {output_path}")


# ==================== 主函数 ====================

def main():
    """主流程"""
    print("=" * 60)
    print("中俄双语历史语料 NLP 分析工具")
    print("=" * 60)

    # 1. 读取文档
    docx_path = 'Data/九一八事变资料.docx'
    print(f"\n[1/6] 读取文档: {docx_path}")
    paragraphs = read_docx(docx_path)
    print(f"  共读取 {len(paragraphs)} 个段落")

    # 2. 分离中俄文
    print("\n[2/6] 分离中俄文文本...")
    chinese_texts, russian_texts, events = separate_languages(paragraphs)
    print(f"  中文段落: {len(chinese_texts)}")
    print(f"  俄文段落: {len(russian_texts)}")
    print(f"  事件分段: {len(events)}")

    # 3. 处理中文
    print("\n[3/6] 处理中文文本（分词、停用词过滤、同义词合并）...")
    chinese_counter = process_chinese(chinese_texts)
    chinese_top20 = get_top_words(chinese_counter, 20)
    chinese_total = sum(chinese_counter.values())
    chinese_norm_freq = calculate_normalized_freq(chinese_counter, chinese_total)
    print(f"  中文有效词数: {chinese_total}")
    print(f"  Top 5: {chinese_top20[:5]}")

    # 4. 处理俄文
    print("\n[4/6] 处理俄文文本（词形还原、词性过滤、停用词过滤）...")
    russian_counter = process_russian(russian_texts)
    russian_top20 = get_top_words(russian_counter, 20)
    russian_total = sum(russian_counter.values())
    russian_norm_freq = calculate_normalized_freq(russian_counter, russian_total)
    print(f"  俄文有效词数: {russian_total}")
    print(f"  Top 5: {russian_top20[:5]}")

    # 5. 生成可视化
    print("\n[5/6] 生成可视化图表...")
    font_path = get_system_font()
    print(f"  使用字体: {font_path}")

    # 中文词云
    generate_wordcloud(
        dict(chinese_counter),
        OUTPUT_DIR / 'wordcloud_chinese.png',
        '中文高频词词云',
        font_path
    )

    # 俄文词云
    generate_wordcloud(
        dict(russian_counter),
        OUTPUT_DIR / 'wordcloud_russian.png',
        '俄文高频词词云',
        font_path
    )

    # 中文柱状图（使用标准化频次）
    chinese_top20_norm = [(w, chinese_norm_freq[w]) for w, _ in chinese_top20]
    generate_bar_chart(
        chinese_top20_norm,
        OUTPUT_DIR / 'barchart_chinese.png',
        '中文 Top 20 高频词（标准化频次）',
        '词汇',
        '标准化频次（每万词）',
        font_path
    )

    # 俄文柱状图（使用标准化频次）
    russian_top20_norm = [(w, russian_norm_freq[w]) for w, _ in russian_top20]
    generate_bar_chart(
        russian_top20_norm,
        OUTPUT_DIR / 'barchart_russian.png',
        '俄文 Top 20 高频词（标准化频次）',
        '词汇',
        '标准化频次（每万词）',
        font_path
    )

    # 6. 导出Excel
    print("\n[6/6] 导出统计数据到Excel...")
    chinese_export = [(w, c, chinese_norm_freq[w]) for w, c in chinese_top20]
    russian_export = [(w, c, russian_norm_freq[w]) for w, c in russian_top20]
    export_to_excel(
        chinese_export,
        russian_export,
        OUTPUT_DIR / 'result.xlsx'
    )

    print("\n" + "=" * 60)
    print("✓ 所有任务完成！")
    print(f"✓ 输出目录: {OUTPUT_DIR.absolute()}")
    print("=" * 60)


if __name__ == '__main__':
    main()
