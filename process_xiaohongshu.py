import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import logging
from urllib.parse import urlparse
import os

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def setup_headers():
    """设置请求头"""
    return {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Cookie': 'abRequestId=3df4775e-8732-5ada-a29a-ce7971b7e706; a1=18eed59d3captrvar6bdsb7kzf8ypujw5g3evitt050000165931; webId=a545d7731bd0f0842664e54ceb094f9f; gid=yYddf2j2iJ28yYddf2jfqk39S0U9Fh0FKDfMq98DWTuiYd28vU7VEh888yK2jqy8KKKJJWD2; web_session=0400698e94013156eb0f0a5512354b13b0f3ae; webBuild=4.47.1; websectiga=984412fef754c018e472127b8effd174be8a5d51061c991aadd200c69a2801d6; sec_poison_id=d7a3a77f-478d-4dd8-9bf6-e0acafaafaf6; acw_tc=0a4ae0b217341421522844429e1e81bf4d0536d9c3ba8ec37594dd05c2afea; unread={%22ub%22:%22675707f2000000000102a0e0%22%2C%22ue%22:%2267597663000000000600cc28%22%2C%22uc%22:27}; xsecappid=xhs-pc-web',  # 替换为你的实际Cookie
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-User': '?1',
    }

def get_note_content(url, headers):
    """
    获取小红书笔记内容
    返回笔记详情和话题标签
    """
    try:
        # 确保URL是完整的
        if not url.startswith('http'):
            url = f'https://www.xiaohongshu.com{url}'
        
        logging.info(f"正在获取笔记内容: {url}")
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        # 设置正确的编码
        response.encoding = 'utf-8'
        
        # 打印响应状态和内容长度，用于调试
        logging.info(f"响应状态码: {response.status_code}")
        logging.info(f"响应内容长度: {len(response.text)}")
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 打印页面标题，确认是否正确获取到页面
        title = soup.title.string if soup.title else "无标题"
        logging.info(f"页面标题: {title}")
        
        # 获取笔记详情
        detail_desc = ''
        detail_element = soup.find(id='detail-desc')
        if detail_element:
            detail_desc = detail_element.get_text(strip=True)
        else:
            logging.warning("未找到detail-desc元素")
            # 尝试其他可能的选择器
            detail_element = soup.select_one('.content')
            if detail_element:
                detail_desc = detail_element.get_text(strip=True)
        
        # 获取话题标签
        hashtags = []
        hashtag_elements = soup.find_all(id='hash-tag')
        if not hashtag_elements:
            # 尝试其他可能的选择器
            hashtag_elements = soup.select('.tag-item')
        
        for element in hashtag_elements:
            hashtags.append(element.get_text(strip=True))
        
        if not detail_desc and not hashtags:
            logging.warning("未能找到笔记详情和话题标签")
            # 打印页面结构以便调试
            logging.debug(f"页面结构: {soup.prettify()[:500]}...")
        
        return detail_desc, ' '.join(hashtags)
    
    except requests.exceptions.RequestException as e:
        logging.error(f"请求失败: {str(e)}")
        return '', ''
    except Exception as e:
        logging.error(f"获取笔记内容失败: {url}, 错误: {str(e)}")
        return '', ''

def process_xiaohongshu_notes(input_file, output_file):
    """
    处理小红书笔记数据
    """
    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        logging.error(f"输入文件不存在: {input_file}")
        return

    try:
        # 读取Excel文件
        logging.info(f"正在读取文件: {input_file}")
        df = pd.read_excel(input_file)
        
        # 检查数据是否为空
        if df.empty:
            logging.error("Excel文件中没有数据")
            return
            
        # 检查是否存在第一列数据
        if len(df.columns) == 0:
            logging.error("Excel文件格式不正确，没有找到URL列")
            return
            
        # 获取请求头
        headers = setup_headers()
        
        # 存储新的数据
        note_details = []
        note_hashtags = []
        
        # 处理每个URL
        for index, url in enumerate(df.iloc[:, 0]):
            if pd.isna(url):  # 检查URL是否为空
                logging.warning(f"第 {index + 1} 行的URL为空，跳过处理")
                note_details.append('')
                note_hashtags.append('')
                continue
                
            logging.info(f"正在处理第 {index + 1}/{len(df)} 条数据")
            
            # 获取笔记内容
            detail, hashtags = get_note_content(url, headers)
            note_details.append(detail)
            note_hashtags.append(hashtags)
            
            # 添加延时，避免请求过于频繁
            time.sleep(2)
        
        # 添加新列
        df['笔记详情'] = note_details
        df['笔记话题'] = note_hashtags
        
        # 保存到新文件
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        logging.info(f"正在保存处理结果到: {output_file}")
        df.to_excel(output_file, index=False)
        logging.info("处理完成！")
        
    except Exception as e:
        logging.error(f"处理过程中发生错误: {str(e)}")

if __name__ == "__main__":
    # 使用相对路径，更简单直接
    input_file = "小红书演示数据使用12月航海.xlsx"
    output_file = "processed_xiaohongshu_notes.xlsx"
    
    # 打印实际使用的文件路径
    logging.info(f"输入文件路径: {input_file}")
    logging.info(f"输出文件路径: {output_file}")
    
    process_xiaohongshu_notes(input_file, output_file) 