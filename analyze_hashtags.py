import pandas as pd
from collections import Counter
import logging

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def analyze_hashtags(input_file, output_file, top_n=50):
    """
    分析话题标签并生成统计报告
    """
    try:
        # 读取Excel文件
        logging.info(f"正在读取文件: {input_file}")
        df = pd.read_excel(input_file)
        
        # 检查是否存在"笔记话题"列
        if '笔记话题' not in df.columns:
            logging.error("未找到'笔记话题'列")
            return
            
        # 存储所有话题
        all_hashtags = []
        
        # 处理每行的话题标签
        for tags in df['笔记话题']:
            if pd.isna(tags):  # 跳过空值
                continue
                
            # 分割话题标签（按空格分割）
            hashtags = [tag.strip() for tag in str(tags).split()]
            # 清理标签（去除#号）
            hashtags = [tag.strip('#') for tag in hashtags if tag.strip()]
            all_hashtags.extend(hashtags)
        
        # 统计话题出现次数
        hashtag_counter = Counter(all_hashtags)
        
        # 获取前N个最常见的话题
        top_hashtags = hashtag_counter.most_common(top_n)
        
        # 写入结果到文件
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(f"Top {top_n} 话题标签统计:\n")
            f.write("=" * 40 + "\n")
            for tag, count in top_hashtags:
                f.write(f"#{tag}: {count}次\n")  # 添加#号显示
        
        # 打印统计信息
        logging.info(f"总共发现 {len(hashtag_counter)} 个不同的话题标签")
        logging.info(f"已将前 {top_n} 个话题写入文件: {output_file}")
        
        # 打印前10个最常见的话题
        logging.info("\n最常见的10个话题：")
        for tag, count in top_hashtags[:10]:
            logging.info(f"#{tag}: {count}次")  # 添加#号显示
            
    except Exception as e:
        logging.error(f"处理过程中发生错误: {str(e)}")
        raise

if __name__ == "__main__":
    input_file = "processed_xiaohongshu_notes.xlsx"
    output_file = "hashtag_statistics.txt"
    
    analyze_hashtags(input_file, output_file) 