import requests
import random
import json
from datetime import datetime
import china_calendar
import os
from docx import Document
import pdfplumber
import fitz  # PyMuPDF

def send_dingtalk_message_by_type(msg_type):
    """
    发送消息到钉钉机器人
    
    Args:
        msg_type (str): 消息类型，"txt" 或 "image"
    """
    # 钉钉机器人的 Webhook 地址
    # webhook_url = "https://oapi.dingtalk.com/robot/send?access_token=79b2d48cb0db1b95df0492cee91192809488c89ce76cfa8549165cc2f7444ce5"
    webhook_url = "https://oapi.dingtalk.com/robot/send?access_token=e65fbbbee753149cdf8a8815300aaff321bff974e83f82b2488e00a37a13a292"
    
    # 从webhook URL中提取access_token
    import urllib.parse
    parsed_url = urllib.parse.urlparse(webhook_url)
    query_params = urllib.parse.parse_qs(parsed_url.query)
    access_token = query_params.get('access_token', [''])[0]
    
    if not access_token:
        print("无法从webhook URL中提取access_token")
        return
    
    # 设置请求头，指定内容类型为 JSON
    headers = {
        "Content-Type": "application/json;charset=utf-8"
    }
    
    try:
        if msg_type == "txt":
            # ---------------------- 共享文件路径配置 ----------------------
            # JSON_FILE_PATH = r"\\192.168.0.3\99-公司共享\API\englishLearn.json"
            JSON_FILE_PATH = r"D:\test\englishLearn.json"
            CET_4_FILE_PATH = r"\\192.168.0.3\99-公司共享\API\CET4\CET4.json"
            WORD_FILE_PATH = r"\\192.168.0.3\99-公司共享\API\word.docx"
            # 获取文件内容
            old_word_data = read_english_learn_json(JSON_FILE_PATH)
            # 读取CET4.json文件，提取随机单词信息
            # cet4_word_info = extract_word_info(CET_4_FILE_PATH)
            # 获取新的单词和词组
            # word = cet4_word_info.get("headWord", "未知单词") + " 【" + cet4_word_info.get("ukphone", "无音标") + "】 " + cet4_word_info.get("pos", "无词性") + ". " + cet4_word_info.get("tran", "无释义")
            # phrase = cet4_word_info.get("phrases第一个pContent", "无词组") + " " + cet4_word_info.get("phrases第一个pCn", "无词组释义")
            # 从Word文件读取单词和词组
            word, phrase = read_word_from_docx(WORD_FILE_PATH)
            # 保存文件
            save_to_shared_json(JSON_FILE_PATH, word, phrase)
            # 构造文本消息
            today = datetime.now().strftime("%Y-%m-%d")
            data = {
                "msgtype": "text",
                "text": {
                    "content": f"每日学英语 {today}\n{old_word_data.get('word', '')}\n{old_word_data.get('phrase', '')}\n{word}\n{phrase}"
                }
            }
            
            # 发送 POST 请求
            response = requests.post(webhook_url, data=json.dumps(data), headers=headers)
        
        elif msg_type == "image":
            PDF_FILE_PATH = r"D:\test\pdf"
            PNG_FILE_PATH = r"D:\test\png"
            
            # 处理PDF文件，获取图片路径
            image_path = process_pdf_files(PDF_FILE_PATH, PNG_FILE_PATH)
            if not image_path:
                print("图片路径不能为空")
                return
            
            # 上传图片到 PicGo
            image_url = upload_image_to_picgo(image_path)
            print(f"图片URL: {image_url}")
            if not image_url:
                print("无法获取图片URL，图片发送失败")
                return
            
            # 构造图片消息
            data = {
                "msgtype": "markdown",
                "markdown": {
                    "title":"每日学英语",
                    "text": f"每日学英语 \n <img src='{image_url}'/>"
                }
            }
            
            # 发送 POST 请求
            response = requests.post(webhook_url, data=json.dumps(data), headers=headers)
        
        else:
            print(f"不支持的消息类型：{msg_type}")
            return
        
        # 打印响应状态码和响应内容，方便排查问题
        print(f"响应状态码: {response.status_code}")
        print(f"响应内容: {response.text}")
        
        # 解析响应结果，判断是否发送成功
        result = response.json()
        if result.get("errcode") == 0:
            print("消息发送成功！")
        else:
            print(f"消息发送失败：{result.get('errmsg')}")
    
    except requests.exceptions.RequestException as e:
        # 捕获请求过程中的异常（如网络错误、连接超时等）
        print(f"请求发送失败，异常信息：{e}")

def send_dingtalk_message():
    today = datetime.now().strftime("%Y-%m-%d")
    is_holiday = china_calendar.is_holiday(today)
    if not is_holiday:
        # 发送文本消息
        # send_dingtalk_message_by_type("txt")
        # 发送图片消息
        send_dingtalk_message_by_type("image")
# ---------------------- 文件保存核心函数 ----------------------
def save_to_shared_json(path, word, phrase):
    """单词和词组保存到共享路径的JSON文件（每次清空原有内容）"""
    # 构造要保存的JSON数据
    save_data = {
        "word": word,
        "phrase": phrase
    }
    
    try:
        # 4. 检查共享路径是否存在，不存在则创建目录
        file_dir = os.path.dirname(path)
        if not os.path.exists(file_dir):
            os.makedirs(file_dir)
        
        # 5. 清空文件并写入新内容（'w'模式会自动清空原有内容）
        with open(path, "w", encoding="utf-8") as f:
            # 格式化JSON，便于阅读
            json.dump(save_data, f, ensure_ascii=False, indent=4)
        
    except PermissionError:
        print(f"权限不足！无法写入文件：{path}")
        print("请检查共享文件夹的读写权限，或确认网络路径是否可访问")
    except FileNotFoundError:
        print(f"路径不存在！请检查：{path}")
    except Exception as e:
        print(f"保存文件异常：{e}")

def read_english_learn_json(path):
    """
    读取共享路径下的englishLearn.json文件内容，并返回JSON数据
    """
    # 1. 检查文件是否存在
    if not os.path.exists(path):
        print(f"错误：文件不存在 → {path}")
        return {}
    
    # 2. 检查是否是文件（而非目录）
    if not os.path.isfile(path):
        print(f"错误：路径不是文件 → {path}")
        return {}
    
    try:
        # 3. 以UTF-8编码读取文件（避免中文乱码）
        with open(path, "r", encoding="utf-8") as f:
            json_data = json.load(f)
            return json_data
    
    except PermissionError:
        print(f"❌ 权限不足：无法读取文件，请检查共享文件夹的访问权限 → {path}")
        return {}
    except json.JSONDecodeError:
        print(f"❌ JSON格式错误：文件内容不是合法的JSON格式 → {path}")
        return {}
    except FileNotFoundError:
        print(f"❌ 文件未找到：确认路径是否正确 → {path}")
        return {}
    except Exception as e:
        print(f"❌ 读取文件异常：{str(e)}")
        return {}

def extract_word_info(file_path):
    """
    读取JSON文件，随机选取包含phrases的headWord，提取对应信息
    若选中单词无phrases，则重新选取，直到找到为止
    
    Args: file_path (str): JSON文件路径
    
    Returns: dict: 包含选中单词的所有目标信息（确保有phrases）
    """
    # 1. 读取并解析JSON文件（标准JSON数组）
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)  # 直接解析合法JSON数组
    except FileNotFoundError:
        return {"error": f"文件 {file_path} 不存在"}
    except json.JSONDecodeError as e:
        return {"error": f"JSON解析失败：{str(e)}"}
    except Exception as e:
        return {"error": f"读取文件出错：{str(e)}"}

    # 2. 验证数据格式
    if not isinstance(data, list) or len(data) == 0:
        return {"error": "JSON数据不是非空列表"}

    # 3. 过滤出包含phrases且phrases非空，且headWord长度小于等于6的单词列表
    valid_words = []
    for word in data:
        # 逐层检查phrases是否存在且非空
        phrase_list = word.get('content', {}) \
                          .get('word', {}) \
                          .get('content', {}) \
                          .get('phrase', {}) \
                          .get('phrases', [])
        # 检查headWord长度是否小于等于6
        head_word = word.get('headWord', '')
        if isinstance(phrase_list, list) and len(phrase_list) > 0 and len(head_word) <= 6:
            valid_words.append(word)

    # 4. 验证有效单词列表（防止所有单词都无phrases或headWord长度超过6）
    if len(valid_words) == 0:
        return {"error": "所有单词都没有有效的phrases数据或headWord长度超过6"}

    # 5. 从有效列表中随机选一个（确保有phrases）
    random_word = random.choice(valid_words)
    head_word = random_word.get('headWord', '未知单词')
    word_content = random_word.get('content', {}).get('word', {}).get('content', {})

    # 6. 提取目标字段（此时phrases一定存在）
    # 提取pos（优先relWord，无则取syno）
    pos_info = []
    # 从relWord提取同根词性
    rel_words = word_content.get('relWord', {}).get('rels', [])
    for rel in rel_words:
        pos = rel.get('pos')
        if pos and pos not in pos_info:
            pos_info.append(pos)
    # 若无relWord，从syno提取词性
    if not pos_info:
        synos = word_content.get('syno', {}).get('synos', [])
        for syn in synos:
            pos = syn.get('pos')
            if pos and pos not in pos_info:
                pos_info.append(pos)
    pos_str = '/'.join(pos_info) if pos_info else '无'

    # 提取核心释义
    tran_info = word_content.get('trans', [{}])[0].get('tranCn', '无')

    # 提取phrases（此时一定有值）
    first_phrase = word_content.get('phrase', {}).get('phrases', [{}])[0]
    p_content = first_phrase.get('pContent', '无')
    p_cn = first_phrase.get('pCn', '无')

    # 7. 整理结果
    result = {
        "headWord": head_word,
        "pos": pos_str,
        "ukphone": word_content.get('ukphone', '无'),
        "tran": tran_info,
        "phrases第一个pContent": p_content,
        "phrases第一个pCn": p_cn
    }

    return result

def read_word_from_docx(file_path):
    """
    读取Word文件，获取一个单词（包含中文发音）和一个词组
    
    Args: file_path (str): Word文件路径
    
    Returns: tuple: (word, phrase) 包含单词和词组的元组
    """
    try:
        # 1. 打开Word文档
        doc = Document(file_path)
        
        # 2. 提取所有段落文本
        all_text = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                all_text.append(text)
        
        # 3. 过滤出符合条件的单词候选（包含/的行及其下一行）
        word_candidates = []
        for i, text in enumerate(all_text):
            if "/" in text and i + 1 < len(all_text):
                # 找到包含/的行，且存在下一行
                word_pair = text + "\n" + all_text[i + 1]
                word_candidates.append(word_pair)
        
        # 4. 过滤出符合条件的词组候选（不包含/和【中文发音】的行）
        phrase_candidates = []
        for text in all_text:
            if "/" not in text and "中文发音" not in text:
                phrase_candidates.append(text)
        
        # 5. 随机选择单词和词组
        word = random.choice(word_candidates) if word_candidates else "无单词"
        phrase = random.choice(phrase_candidates) if phrase_candidates else "无词组"
        
        return word, phrase
        
    except FileNotFoundError:
        print(f"错误：文件不存在 → {file_path}")
        return "无单词", "无词组"
    except Exception as e:
        print(f"读取Word文件异常：{str(e)}")
        return "无单词", "无词组"

def get_dingtalk_access_token(app_key, app_secret):
    """
    获取钉钉企业access_token
    
    Args:
        app_key (str): 应用的AppKey
        app_secret (str): 应用的AppSecret
    
    Returns:
        str: 成功返回access_token，失败返回None
    """
    try:
        url = f"https://oapi.dingtalk.com/gettoken?appkey={app_key}&appsecret={app_secret}"
        response = requests.get(url)
        result = response.json()
        
        if result.get('errcode') == 0:
            return result.get('access_token')
        else:
            print(f"获取access_token失败：{result.get('errmsg')}")
            return None
    except Exception as e:
        print(f"获取access_token异常：{str(e)}")
        return None

def upload_image_to_dingtalk(image_path, access_token):
    """
    上传图片到钉钉服务器并返回media_id
    
    Args:
        image_path (str): 图片文件路径
        access_token (str): 钉钉企业的access_token
    
    Returns:
        str: 上传成功返回media_id，失败返回None
    """
    try:
        # 钉钉上传文件接口
        upload_url = f"https://oapi.dingtalk.com/media/upload?access_token={access_token}&type=image"
        
        # 读取图片文件
        with open(image_path, 'rb') as f:
            files = {'media': f}
            response = requests.post(upload_url, files=files)
        
        # 解析响应
        result = response.json()
        if result.get('errcode') == 0:
            return result.get('media_id')
        else:
            print(f"上传图片失败：{result.get('errmsg')}")
            return None
    except Exception as e:
        print(f"上传图片异常：{str(e)}")
        return None

def upload_image_to_picgo(image_path):
    """
    上传图片到 PicGo 并返回图片 URL
    
    Args:
        image_path (str): 图片文件路径
    
    Returns:
        str: 上传成功返回图片 URL，失败返回None
    """
    try:
        # PicGo 上传接口
        upload_url = "https://www.picgo.net/api/1/upload"
        # API 密钥
        apikey = "chv_S6G4f_0b30698abdccad5ba52b292ebfd5a4f1872526551fc5296307e0a5b478d64264_f34e296f0f18d5666f85d19dfd0850880c07bc24ccd0ec206fab43b3604e0199"
        
        # 读取图片文件
        with open(image_path, 'rb') as f:
            # POST 请求参数中包含 key 和 source 字段
            files = {
                'source': (os.path.basename(image_path), f, 'image/png')
            }
            data = {
                'key': apikey
            }
            headers = {
                'Accept': 'application/json'
            }
            # 不需要手动设置 Content-Type，requests 会自动处理 multipart/form-data
            response = requests.post(upload_url, files=files, data=data, headers=headers)
        
        # 解析响应
        result = response.json()
        if result.get('status_code') == 200:
            return result.get('image', {}).get('url')
        else:
            print(f"上传图片到 PicGo 失败：{result.get('status_txt')}")
            return None
    except Exception as e:
        print(f"上传图片到 PicGo 异常：{str(e)}")
        return None

def process_pdf_files(pdf_folder, png_folder):
    """
    处理PDF文件，找到包含今天日期的页面并转换为PNG图片
    
    Args:
        pdf_folder (str): PDF文件所在目录
        png_folder (str): PNG图片保存目录
    
    Returns:
        str: 成功返回图片路径，失败返回None
    """
    try:
        # 1. 获取今天的日期格式（如3/20）
        today = datetime.now()
        date_str = f"{today.month}/{today.day}"
        
        # 2. 确保PNG保存目录存在
        if not os.path.exists(png_folder):
            os.makedirs(png_folder)
        
        # 3. 遍历PDF文件夹中的所有PDF文件
        pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            print(f"未找到PDF文件：{pdf_folder}")
            return
        
        # 4. 处理每个PDF文件
        for pdf_file in pdf_files:
            pdf_path = os.path.join(pdf_folder, pdf_file)
            print(f"处理文件：{pdf_file}")
            
            # 5. 打开PDF文件查找包含日期的页面
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    # 提取页面文本
                    text = page.extract_text() or ""
                    
                    # 检查是否包含今天的日期
                    if date_str in text:
                        print(f"找到包含日期 {date_str} 的页面：{page_num}")
                        
                        # 6. 使用PyMuPDF将页面转换为PNG图片
                        doc = fitz.open(pdf_path)
                        page = doc.load_page(page_num - 1)  # PyMuPDF页面索引从0开始
                        pix = page.get_pixmap()
                        
                        # 7. 保存图片到指定的PNG文件夹
                        png_filename = f"{date_str.replace('/', '-')}.png"  # 替换/为-以避免路径问题
                        png_path = os.path.join(png_folder, png_filename)
                        pix.save(png_path)
                        print(f"保存图片：{png_path}")
                        doc.close()
                        return png_path  # 返回图片路径
        
        print(f"未找到包含日期 {date_str} 的页面")
        return None
        
    except FileNotFoundError:
        print(f"错误：PDF目录不存在 → {pdf_folder}")
        return None
    except Exception as e:
        print(f"处理PDF文件异常：{str(e)}")
        return None

if __name__ == "__main__":
    send_dingtalk_message()