import requests
import random
import json
from datetime import datetime
import china_calendar
import os

def send_dingtalk_message():
    today = datetime.now().strftime("%Y-%m-%d")
    is_holiday = china_calendar.is_holiday(today)
    if not is_holiday:
        # ---------------------- 共享文件路径配置 ----------------------
        JSON_FILE_PATH = r"\\192.168.0.3\99-公司共享\API\englishLearn.json"
        CET_4_FILE_PATH = r"\\192.168.0.3\99-公司共享\API\CET4\CET4.json"
        # 获取文件内容
        old_word_data = read_english_learn_json(JSON_FILE_PATH)
        # 读取CET4.json文件，提取随机单词信息
        cet4_word_info = extract_word_info(CET_4_FILE_PATH)
        # 获取新的单词和词组
        word = cet4_word_info.get("headWord", "未知单词") + " 【" + cet4_word_info.get("ukphone", "无音标") + "】 " + cet4_word_info.get("pos", "无词性") + ". " + cet4_word_info.get("tran", "无释义")
        phrase = cet4_word_info.get("phrases第一个pContent", "无词组") + " " + cet4_word_info.get("phrases第一个pCn", "无词组释义")
        # 保存文件
        save_to_shared_json(JSON_FILE_PATH, word, phrase)
        # 钉钉机器人的 Webhook 地址
        # url = "https://oapi.dingtalk.com/robot/send?access_token=79b2d48cb0db1b95df0492cee91192809488c89ce76cfa8549165cc2f7444ce5"
        url = "https://oapi.dingtalk.com/robot/send?access_token=e65fbbbee753149cdf8a8815300aaff321bff974e83f82b2488e00a37a13a292"
        
        # 构造符合钉钉机器人要求的请求体
        # 钉钉机器人要求消息体必须是 JSON 格式，且包含 msgtype 和 text 字段
        data = {
            "msgtype": "text",
            "text": {
                "content": f"每日学英语 {today}\n{old_word_data.get('word', '')}\n{old_word_data.get('phrase', '')}\n{word}\n{phrase}"  # 你要发送的测试内容
            }
        }
        
        # 设置请求头，指定内容类型为 JSON
        headers = {
            "Content-Type": "application/json;charset=utf-8"
        }
        
        try:
            # 发送 POST 请求
            # 将字典格式的 data 转换为 JSON 字符串
            response = requests.post(url, data=json.dumps(data), headers=headers)
            
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

    # 3. 过滤出包含phrases且phrases非空的单词列表
    valid_words = []
    for word in data:
        # 逐层检查phrases是否存在且非空
        phrase_list = word.get('content', {}) \
                          .get('word', {}) \
                          .get('content', {}) \
                          .get('phrase', {}) \
                          .get('phrases', [])
        if isinstance(phrase_list, list) and len(phrase_list) > 0:
            valid_words.append(word)
    
    # 4. 验证有效单词列表（防止所有单词都无phrases）
    if len(valid_words) == 0:
        return {"error": "所有单词都没有有效的phrases数据"}

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

if __name__ == "__main__":
    send_dingtalk_message()