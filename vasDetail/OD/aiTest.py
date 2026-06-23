import requests

# ========== 你只需要改这里 ==========
APPKEY = "dingrww58ramug37ijoh"
APPSECRET = "XhlOHGtGI_dTIdHC_9bVJCgyzCXvlHKiKKwGuYce53BUzeFf1KBfrKYgElpFxMZ1"
# ===================================

# 1. 获取access_token（老接口没变）
def get_token():
    url = f"https://oapi.dingtalk.com/gettoken?appkey={APPKEY}&appsecret={APPSECRET}"
    return requests.get(url).json()["access_token"]

# 2. 获取所有流程（还是老接口，没变）
def get_all_processes(token):
    url = f"https://oapi.dingtalk.com/topapi/process/listbyuserid?access_token={token}"
    return requests.get(url).json().get("process_list", [])

# 3. 新版：用 Schema 接口判断是否有附件
def has_attachment(token, process_code):
    url = f"https://api.dingtalk.com/v1.0/workflow/forms/schemas?processCode={process_code}"
    headers = {"x-acs-dingtalk-access-token": token}
    res = requests.get(url, headers=headers).json()

    # 遍历 schema 找 attachment 组件
    schema = res.get("schema", {})
    components = schema.get("components", [])
    print(components)
    for comp in components:
        if comp.get("type") == "DDAttachment":
            return True
    return False

# 主程序
if __name__ == "__main__":
    token = get_token()
    processes = get_all_processes(token)

    print("=== 含附件的审批流程（新版Schema接口）===")
    for p in processes:
        name = p.get("name")
        code = p.get("process_code")
        if has_attachment(token, code):
            print(f"✅ {name} | {code}")