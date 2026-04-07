#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
字体检测辅助脚本
用于检测系统中可用的中文字体，帮助配置 analyze_text.py
"""

import os
import platform
import matplotlib.font_manager as fm

def find_chinese_fonts():
    """查找系统中可用的中文字体"""
    print("="*60)
    print("正在扫描系统字体...")
    print("="*60)

    # 获取所有字体
    fonts = fm.findSystemFonts()
    chinese_fonts = []

    # 常见中文字体关键词
    chinese_keywords = [
        'simhei', 'simsun', 'simkai', 'simfang',  # Windows
        'msyh', 'microsoft',  # 微软雅黑
        'pingfang', 'heiti', 'songti', 'kaiti',  # Mac
        'wqy', 'noto', 'droid', 'arphic'  # Linux
    ]

    for font_path in fonts:
        font_name_lower = os.path.basename(font_path).lower()
        if any(keyword in font_name_lower for keyword in chinese_keywords):
            chinese_fonts.append(font_path)

    return chinese_fonts


def get_recommended_font():
    """根据操作系统推荐字体"""
    system = platform.system()

    recommendations = {
        'Windows': [
            'C:/Windows/Fonts/simhei.ttf',  # 黑体
            'C:/Windows/Fonts/msyh.ttc',    # 微软雅黑
            'C:/Windows/Fonts/simsun.ttc',  # 宋体
        ],
        'Darwin': [  # Mac
            '/System/Library/Fonts/PingFang.ttc',
            '/Library/Fonts/Arial Unicode.ttf',
            '/System/Library/Fonts/STHeiti Light.ttc',
        ],
        'Linux': [
            '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
            '/usr/share/fonts/truetype/arphic/uming.ttc',
            '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        ]
    }

    return recommendations.get(system, [])


def test_font(font_path):
    """测试字体是否可用"""
    try:
        from matplotlib import font_manager
        font_prop = font_manager.FontProperties(fname=font_path)
        return True
    except:
        return False


def main():
    print("\n字体检测工具")
    print("="*60)

    # 1. 显示操作系统信息
    system = platform.system()
    print(f"\n当前操作系统: {system}")

    # 2. 显示推荐字体
    print(f"\n推荐字体路径（按优先级排序）:")
    print("-"*60)
    recommended = get_recommended_font()
    for i, font_path in enumerate(recommended, 1):
        exists = "✓ 存在" if os.path.exists(font_path) else "✗ 不存在"
        print(f"{i}. {font_path}")
        print(f"   状态: {exists}")

    # 3. 扫描系统中的中文字体
    print(f"\n扫描到的中文字体:")
    print("-"*60)
    chinese_fonts = find_chinese_fonts()

    if chinese_fonts:
        for i, font_path in enumerate(chinese_fonts[:10], 1):  # 只显示前10个
            print(f"{i}. {font_path}")
        if len(chinese_fonts) > 10:
            print(f"... 还有 {len(chinese_fonts) - 10} 个字体未显示")
    else:
        print("未找到中文字体")

    # 4. 给出配置建议
    print(f"\n配置建议:")
    print("="*60)

    available_font = None
    for font_path in recommended:
        if os.path.exists(font_path):
            available_font = font_path
            break

    if not available_font and chinese_fonts:
        available_font = chinese_fonts[0]

    if available_font:
        print(f"✓ 建议在 analyze_text.py 中设置:")
        print(f"\n  FONT_PATH = '{available_font}'")
        print(f"\n或使用相对路径（如果字体在项目目录）:")
        print(f"\n  FONT_PATH = '{os.path.basename(available_font)}'")
    else:
        print("✗ 未找到可用的中文字体！")
        print("\n解决方案:")
        print("1. 下载中文字体文件（如 simhei.ttf）")
        print("2. 将字体文件放到项目目录")
        print("3. 设置 FONT_PATH = 'simhei.ttf'")

    # 5. 测试matplotlib配置
    print(f"\n测试matplotlib字体配置:")
    print("-"*60)
    try:
        import matplotlib.pyplot as plt
        plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
        plt.rcParams['axes.unicode_minus'] = False
        print("✓ matplotlib中文支持配置成功")
    except Exception as e:
        print(f"✗ matplotlib配置失败: {e}")

    print("\n" + "="*60)
    print("检测完成！")
    print("="*60 + "\n")


if __name__ == '__main__':
    main()
