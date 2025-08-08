from flask import Flask, jsonify, request, send_file
import os
import uuid
import json
from datetime import datetime
import pymssql

app = Flask(__name__)

# 配置
PHOTOS_DIR = "D:/photos"  # 照片存储目录
VOTES_FILE = "D:/votes.json"  # 投票结果存储文件
DEVICE_ID_COOKIE = "vote_device_id"  # 设备ID的cookie名称
# sql服务器名
SERVER_NAME = '192.168.0.11'
# 登陆用户名和密码
USER_NAME = 'sa'
PASSWORD = 'jiangbin@007'
# 数据库名
DB_NAME = 'ESApp1'

def get_device_id():
    """获取或生成设备唯一ID"""
    device_id = request.cookies.get(DEVICE_ID_COOKIE)
    
    if not device_id:
        # 生成基于IP和用户代理的设备ID
        ip = request.remote_addr
        user_agent = request.headers.get('User-Agent', '')
        device_id = str(uuid.uuid5(uuid.NAMESPACE_DNS, f"{ip}-{user_agent}"))
    
    return device_id

def has_voted(device_id):
    """检查设备是否已投票"""
    try:
        # 建立连接并获取PO数据
        conn = pymssql.connect(SERVER_NAME, USER_NAME, PASSWORD, DB_NAME)
        cursor = conn.cursor()
        select_sql = 'select distinct device_id from D_Votes WHERE device_id = \'{0}\''.format(device_id)
        cursor.execute(select_sql)
        row = cursor.fetchall()
        # 判断row是否为空
        if not row:
            return False
        return device_id
    except:
        return False
    finally:
        cursor.close()
        conn.close()

@app.route('/api/check_vote', methods=['GET'])
def check_vote():
    """检查当前设备是否已投票"""
    try:
        device_id = get_device_id()
        return jsonify({
            "success": True,
            "has_voted": has_voted(device_id)
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

def save_vote(device_id, photo_name):
    """保存投票结果"""
    try:
        # 建立连接并获取cursor
        conn = pymssql.connect(SERVER_NAME, USER_NAME, PASSWORD, DB_NAME)
        cursor = conn.cursor()
        values = (device_id, photo_name, str(datetime.now()).split('.')[0])
        insertSql = 'INSERT INTO D_Votes (device_id, photo, timestamp) VALUES (%s, %s, %s)'
        cursor.execute(insertSql, values)
        conn.commit()
        conn.close()
    except:
        data = {"votes": [], "devices": {}}
    
    return True

@app.route('/api/photos', methods=['GET'])
def get_photos():
    """获取照片列表"""
    try:
        # 获取D盘中的所有照片文件
        photos = []
        id = 0
        for file in os.listdir(PHOTOS_DIR):
            if file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
                id += 1
                photos.append({
                    "id": id,
                    "name": os.path.splitext(file)[0],
                    "filename": file
                })
        return jsonify({"success": True, "photos": photos})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/photo/<filename>', methods=['GET'])
def get_photo(filename):
    """获取单个照片文件"""
    try:
        # 防止路径遍历攻击
        if '..' in filename or filename.startswith('/'):
            return jsonify({"success": False, "error": "Invalid filename"}), 400
            
        photo_path = os.path.join(PHOTOS_DIR, filename)
        
        if not os.path.exists(photo_path):
            return jsonify({"success": False, "error": "Photo not found"}), 404
            
        return send_file(photo_path)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/vote', methods=['POST'])
def vote():
    """提交投票"""
    try:
        data = request.get_json()
        photo_name = data.get('photo')
        
        if not photo_name:
            return jsonify({"success": False, "error": "Missing photo name"}), 400
        
        device_id = get_device_id()
        
        # 检查是否已投票
        if has_voted(device_id):
            return jsonify({"success": False, "error": "Already voted"}), 400
        
        # 保存投票
        if save_vote(device_id, photo_name):
            response = jsonify({"success": True, "message": "Vote recorded"})
            # 设置设备ID cookie (有效期30天)
            response.set_cookie(DEVICE_ID_COOKIE, device_id, max_age=30*24*60*60)
            return response
        else:
            return jsonify({"success": False, "error": "Failed to save vote"}), 500
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/results', methods=['GET'])
def get_results():
    """获取投票结果"""
    try:
        with open(VOTES_FILE, "r") as f:
            data = json.load(f)
            return jsonify({"success": True, "results": data["votes"]})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)