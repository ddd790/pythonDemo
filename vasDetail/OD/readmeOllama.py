import os
import re
from langchain.llms import Ollama
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain

# ===================== 【你们ERP配置】 =====================
JAVA_ERP_SRC_PATH = r"E:\motives-erp\motives-erp\motives-erp-service\src\main\java"  # 改成你们Java源码路径
OUTPUT_FOLDER = r"E:\ERP操作流程手册"
# ==========================================================

# 创建输出目录
if not os.path.exists(OUTPUT_FOLDER):
    os.mkdir(OUTPUT_FOLDER)

# =============== 【使用最新的 qwen3.5:9b】 ===============
llm = Ollama(model="qwen3.5:9b", temperature=0.05)  # 温度越低，流程越准确
# ==========================================================

# 超强Prompt（专门优化给 qwen3.5:9b）
prompt = PromptTemplate(
    input_variables=["code"],
    template="""
你是专业的ERP业务流程分析师。
请根据下面的Java后端代码，**自动生成员工可直接使用的标准操作流程（SOP）**。

要求：
1. 功能名称
2. 适用角色
3. 功能说明
4. 操作步骤（1、2、3、4、5，简单易懂）
5. 状态流转逻辑
6. 注意事项

不要代码术语，输出给普通员工使用的操作手册。

==================== 代码 ====================
{code}
==================================================
""",
)

# AI 执行链
chain = LLMChain(llm=llm, prompt=prompt)

# 扫描Java文件
def scan_java_files(root):
    java_files = []
    for dirpath, _, filenames in os.walk(root):
        for f in filenames:
            if f.endswith(".java"):
                java_files.append(os.path.join(dirpath, f))
    return java_files

# 只处理业务核心文件（Controller/Service/实体类）
def is_business_code(content):
    return any(key in content for key in ["Service", "Controller", "ServiceImpl", "save", "update", "status", "audit", "approve"])

# 生成操作流程
def generate_sop(file_path, content):
    try:
        print(f"正在生成 → {os.path.basename(file_path)}")
        sop = chain.run(content)
        out_file = os.path.join(OUTPUT_FOLDER, os.path.basename(file_path).replace(".java", ".md"))
        with open(out_file, "w", encoding="utf-8") as f:
            f.write(f"# {os.path.basename(file_path)}\n\n")
            f.write(sop)
        print(f"✅ 已生成：{out_file}\n")
    except Exception as e:
        print(f"❌ 生成失败：{e}")

# 主程序
if __name__ == "__main__":
    print("===== Java ERP → 操作流程自动生成（qwen3.5:9b）=====")
    java_files = scan_java_files(JAVA_ERP_SRC_PATH)
    print(f"扫描到Java文件：{len(java_files)}")

    for file in java_files:
        try:
            with open(file, "r", encoding="utf-8") as f:
                content = f.read()
                if is_business_code(content):
                    generate_sop(file, content)
        except:
            pass

    print("\n🎉 全部操作流程生成完成！文件夹：" + OUTPUT_FOLDER)