import pandas as pd
import jieba
import jieba.analyse
from collections import defaultdict
import logging
import os
import time

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def extract_keywords(title, topK=3):
    """
    从标题中提取关键词
    topK: 每个标题提取的关键词数量
    """
    # 使用 jieba.analyse.extract_tags 提取关键词，允许自定义停用词
    keywords = jieba.analyse.extract_tags(title, topK=topK, allowPOS=('n', 'v', 'a'))  # 只提取名词、动词、形容词
    return keywords

def analyze_titles(input_file, output_file):
    """
    分析标题关键词并生成统计报告
    """
    try:
        # 读取Excel文件
        logging.info(f"正在读取文件: {input_file}")
        df = pd.read_excel(input_file)
        
        # 检查是否存在"笔记标题"列
        if '笔记标题' not in df.columns:
            logging.error("未找到'笔记标题'列")
            return
            
        # 用于存储关键词统计
        keyword_stats = defaultdict(lambda: {'count': 0, 'titles': []})
        
        # 处理每个标题
        total_titles = len(df)
        for index, row in df.iterrows():
            title = str(row['笔记标题'])
            logging.info(f"正在处理第 {index + 1}/{total_titles} 个标题: {title}")
            
            # 提取关键词
            keywords = extract_keywords(title)
            
            # 更新统计信息
            for keyword in keywords:
                keyword_stats[keyword]['count'] += 1
                keyword_stats[keyword]['titles'].append(title)
        
        # 转换为DataFrame并排序
        stats_list = []
        for keyword, data in keyword_stats.items():
            stats_list.append({
                '关键词': keyword,
                '出现次数': data['count'],
                '相关标题': '\n'.join(set(data['titles']))  # 使用set去重
            })
        
        result_df = pd.DataFrame(stats_list)
        result_df = result_df.sort_values('出现次数', ascending=False)
        
        # 保存结果前先检查文件是否被占用
        try:
            # 尝试使用新的文件名
            base_name = os.path.splitext(output_file)[0]
            ext = os.path.splitext(output_file)[1]
            counter = 1
            while os.path.exists(output_file):
                output_file = f"{base_name}_{counter}{ext}"
                counter += 1
            
            logging.info(f"正在保存分析结果到: {output_file}")
            result_df.to_excel(output_file, index=False)
            logging.info("分析完成！")
        except PermissionError:
            logging.error(f"无法保存到 {output_file}，文件可能被占用")
            # 尝试保存到临时文件
            temp_file = f"temp_{int(time.time())}_{os.path.basename(output_file)}"
            logging.info(f"尝试保存到临时文件: {temp_file}")
            result_df.to_excel(temp_file, index=False)
            logging.info(f"结果已保存到临时文件: {temp_file}")
        
        # 打印前10个高频关键词
        logging.info("\n最常见的10个关键词：")
        for _, row in result_df.head(10).iterrows():
            logging.info(f"{row['关键词']}: {row['出现次数']}次")
            
    except Exception as e:
        logging.error(f"处理过程中发生错误: {str(e)}")
        raise

def setup_custom_dictionary():
    """
    设置自定义词典
    """
    custom_words = [
        # 平台相关
        '小红书', '探店', '种草', '拔草', '好物分享',
        # 常用修饰词
        '必看', '推荐', '分享', '安利', '测评',
        # 内容类型
        '教程', '攻略', '指南', '合集', '清单',
        # 情感表达
        '超级', '绝绝子', '无敌', '绝美', '太爱了',
        # 场景词
        '上班族', '学生党', '新手', '达人',
        # 其他常见组合
        '一键', '速成', '快速', '简单', '效果好'
    ]
    
    # 可以为词语指定词频和词性
    for word in custom_words:
        jieba.add_word(word, freq=1000)  # 设置较高的词频确保被优先识别
    
    # 创建停用词文件
    stop_words = ['的', '了', '啦', '呢', '哦', '吧', '啊', '和', '与', '及', '或', '等', '这', '那', '是']
    stop_words_file = 'stop_words.txt'
    
    # 写入停用词文件
    with open(stop_words_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(stop_words))
    
    # 设置停用词文件
    jieba.analyse.set_stop_words(stop_words_file)

if __name__ == "__main__":
    input_file = "processed_xiaohongshu_notes.xlsx"
    output_file = "title_keywords_analysis.xlsx"
    
    # 确保输出文件没有被占用
    if os.path.exists(output_file):
        try:
            # 尝试删除已存在的文件
            os.remove(output_file)
        except PermissionError:
            logging.warning(f"无法删除已存在的文件 {output_file}，将使用新的文件名")
    
    # 设置自定义词典
    setup_custom_dictionary()
    
    analyze_titles(input_file, output_file) 