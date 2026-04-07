# 九一八事变历史语料NLP分析系统

## 📋 项目简介

本项目是一套专业的Python文本分析工具，用于处理中俄双语历史文献，实现自动化的词频统计、文本可视化和数据导出功能。

## 🚀 快速开始

### 1. 安装依赖

```bash
# 方式1：使用requirements.txt（推荐）
pip install -r requirements.txt

# 方式2：手动安装
pip install python-docx jieba wordcloud matplotlib pandas openpyxl pymorphy3 pymorphy3-dicts-ru nltk
```

### 2. 下载NLTK数据

首次运行时，脚本会自动下载俄语停用词数据。如果自动下载失败，可手动执行：

```python
import nltk
nltk.download('stopwords')
```

### 3. 配置字体（防止中文乱码）

脚本默认使用 `simhei.ttf`（黑体）。如果遇到字体问题，请修改 `analyze_text.py` 中的 `FONT_PATH` 变量：

**Windows系统：**
```python
FONT_PATH = 'C:/Windows/Fonts/simhei.ttf'  # 黑体
# 或
FONT_PATH = 'C:/Windows/Fonts/msyh.ttc'    # 微软雅黑
```

**Mac系统：**
```python
FONT_PATH = '/System/Library/Fonts/PingFang.ttc'
```

**Linux系统：**
```python
FONT_PATH = '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc'
```

### 4. 运行分析

```bash
python analyze_text.py
```

## 📁 项目结构

```
wordCloud/
├── Data/
│   └── 九一八事变资料.docx    # 输入文档
├── output/                     # 输出目录（自动创建）
│   ├── wordcloud_chinese.png   # 中文词云图
│   ├── wordcloud_russian.png   # 俄文词云图
│   ├── barchart_chinese.png    # 中文柱状图
│   ├── barchart_russian.png    # 俄文柱状图
│   └── result.xlsx             # 统计数据Excel
├── analyze_text.py             # 主分析脚本
├── requirements.txt            # 依赖包列表
└── README_使用说明.md          # 本文件
```

## 🔧 核心功能

### 1. 文本自动分离
- 自动识别并分离中文和俄文段落
- 基于日期标记进行事件划分
- 支持多种日期格式识别

### 2. 中文文本处理
- **分词**：使用jieba进行精准中文分词
- **停用词过滤**：内置常见停用词表
- **同义词合并**：支持自定义同义词映射

### 3. 俄语文本处理（专业级）
- **词形还原**：使用pymorphy3将不同词形统一为基本形式
  - 例：война, войны, войну → война
- **词性过滤**：只保留名词、动词、形容词、副词等实词
- **停用词过滤**：基于NLTK俄语停用词库
- **同义词合并**：支持自定义俄语同义词映射

### 4. 频次统计
- **绝对频次**：统计词汇在整个语料中的出现次数
- **标准化频次**：计算每万词频次，消除不同事件篇幅差异的影响
  - 公式：`(词频 / 事件总词数) × 10000`

### 5. 数据可视化
- **词云图**：直观展示高频词分布
- **柱状图**：使用标准化频次，便于跨事件比较
- **防乱码**：自动配置中俄文字体

### 6. 数据导出
- **Excel多表格**：
  - `高频词统计`：中俄文合并数据
  - `中文统计`：中文Top20详细数据
  - `俄文统计`：俄文Top20详细数据
  - `事件划分`：按日期划分的事件信息
- **预留字段**：词性、语义分类（可手动标注）

## ⚙️ 自定义配置

### 修改同义词映射

在 `analyze_text.py` 中找到以下全局变量并添加映射：

```python
# 中文同义词
CHN_SYNONYMS = {
    '日本军': '日军',
    '日本人': '日本',
    '东北军': '中国军队',
    # 在此添加更多...
}

# 俄语同义词
RUS_SYNONYMS = {
    'японцы': 'японский',
    'японец': 'японский',
    # 在此添加更多...
}
```

### 修改停用词

```python
# 添加中文停用词
CHN_STOPWORDS.update({'新增词1', '新增词2'})

# 添加俄语停用词
RUS_STOPWORDS.update({'новое_слово1', 'новое_слово2'})
```

### 调整Top N数量

修改主函数中的 `top_n` 参数：

```python
chinese_freq_df = calculate_frequencies(chinese_words, top_n=30)  # 改为Top30
```

## 📊 输出说明

### 1. 词云图
- 文件名：`wordcloud_chinese.png`, `wordcloud_russian.png`
- 尺寸：1200×800像素，300 DPI
- 最多显示100个词汇

### 2. 柱状图
- 文件名：`barchart_chinese.png`, `barchart_russian.png`
- 纵轴：标准化频次（每万词）
- 横轴：Top20高频词
- 柱子上方标注具体数值

### 3. Excel统计表
- 文件名：`result.xlsx`
- 包含字段：
  - 语言：中文/俄文
  - 排名：1-20
  - 词汇：词条内容
  - 词性：预留字段（可手动填写）
  - 语义分类：预留字段（如主体类、行动类等）
  - 绝对频次：总出现次数
  - 标准化频次：每万词频次

## 🐛 常见问题

### Q1: 提示找不到字体文件
**解决方案：**
1. 检查系统字体目录
2. 修改 `FONT_PATH` 为完整路径
3. 或下载字体文件放到项目目录

### Q2: 俄语词形还原不准确
**解决方案：**
1. 确认已安装 `pymorphy3-dicts-ru`
2. 检查俄语文本编码是否正确
3. 可在 `RUS_SYNONYMS` 中手动添加映射

### Q3: 中文分词效果不理想
**解决方案：**
1. 使用 `jieba.add_word('专有名词')` 添加自定义词典
2. 调整停用词表
3. 在同义词字典中合并相似词

### Q4: 内存不足
**解决方案：**
1. 减少 `top_n` 数量
2. 分批处理大文件
3. 降低词云图分辨率

## 📝 技术栈

- **文档处理**：python-docx
- **中文NLP**：jieba
- **俄语NLP**：pymorphy3, nltk
- **数据分析**：pandas
- **可视化**：matplotlib, wordcloud
- **数据导出**：openpyxl

## 📄 许可证

本项目仅供学术研究使用。

## 👨‍💻 维护者

数据分析专家 | 2026-04-07

---

**祝分析顺利！如有问题，请检查上述配置或查看代码注释。**
