import re
import os

def merge_and_clean_stopwords(file1, file2, output_file):
    # 正则表达式：匹配任何包含半角 (a-z, A-Z) 或 全角 (ａ-ｚ, Ａ-Ｚ) 英文字母的字符串
    # 只要包含英文字母就会被过滤掉
    english_pattern = re.compile(r'[a-zA-Zａ-ｚＡ-Ｚ]')
    
    merged_stopwords = set()
    
    # 遍历两个文件
    for file_path in [file1, file2]:
        if not os.path.exists(file_path):
            print(f"⚠️ 警告：找不到文件 {file_path}，请确保文件与脚本在同一目录下。")
            continue
            
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                # 去除两端空白字符（包括换行符）
                word = line.strip()
                
                # 跳过空行
                if not word:
                    continue
                
                # 如果字符串中不包含英文字母，则加入集合
                if not english_pattern.search(word):
                    merged_stopwords.add(word)

    # 将集合转换为列表，并进行排序（方便查看，按照 Unicode 编码排序）
    sorted_stopwords = sorted(list(merged_stopwords))
    
    # 将结果写入新的合并文件
    with open(output_file, 'w', encoding='utf-8') as f:
        for word in sorted_stopwords:
            f.write(word + '\n')
            
    print(f"✅ 处理完成！")
    print(f"提取后共保留了 {len(sorted_stopwords)} 个非英文的唯一停用词。")
    print(f"结果已保存至: {output_file}")

if __name__ == "__main__":
    # 请确保这两个 txt 文件与本 python 脚本放在同一个文件夹下
    input_file1 = 'stopwords_cn.txt'
    input_file2 = 'stopwords_cn2.txt'
    output_file = 'stopwords_cn_all.txt'
    
    merge_and_clean_stopwords(input_file1, input_file2, output_file)