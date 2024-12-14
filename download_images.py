import pandas as pd
import requests
import os
import logging
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor
import time

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def download_image(url, save_dir):
    """
    下载单个图片
    """
    try:
        # 从URL中提取文件名
        filename = os.path.basename(urlparse(url).path)
        if not filename.endswith(('.jpg', '.jpeg', '.png', '.gif')):
            filename += '.jpg'  # 默认使用jpg格式
            
        # 构建保存路径
        save_path = os.path.join(save_dir, filename)
        
        # 检查文件是否已存在
        if os.path.exists(save_path):
            logging.info(f"图片已存在，跳过: {filename}")
            return True
            
        # 发送请求下载图片
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        # 保存图片
        with open(save_path, 'wb') as f:
            f.write(response.content)
            
        logging.info(f"成功下载图片: {filename}")
        return True
        
    except Exception as e:
        logging.error(f"下载图片失败 {url}: {str(e)}")
        return False

def process_images(excel_file):
    """
    处理Excel文件并下载符合条件的图片
    """
    try:
        # 读取Excel文件
        logging.info(f"正在读取文件: {excel_file}")
        df = pd.read_excel(excel_file)
        
        # 检查必要的列是否存在
        required_columns = ['封面地址', '粉丝数', '互动量']
        for col in required_columns:
            if col not in df.columns:
                logging.error(f"未找到必要的列: {col}")
                return
        
        # 创建保存图片的目录
        save_dir = 'downloaded_images'
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
            
        # 筛选符合条件的数据
        filtered_df = df[
            (df['粉丝数'] < 1000) & 
            (df['互动量'] > 100)
        ]
        
        if filtered_df.empty:
            logging.info("没有找到符合条件的数据")
            return
            
        logging.info(f"找到 {len(filtered_df)} 条符合条件的数据")
        
        # 使用线程池下载图片
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = []
            for url in filtered_df['封面地址']:
                if pd.isna(url):  # 跳过空值
                    continue
                futures.append(
                    executor.submit(download_image, url, save_dir)
                )
                time.sleep(0.5)  # 添加延时避免请求过快
            
            # 等待所有下载完成
            success_count = sum(1 for future in futures if future.result())
            
        logging.info(f"下载完成! 成功: {success_count}, 总数: {len(futures)}")
        
    except Exception as e:
        logging.error(f"处理过程中发生错误: {str(e)}")

if __name__ == "__main__":
    excel_file = "processed_xiaohongshu_notes.xlsx"
    process_images(excel_file) 