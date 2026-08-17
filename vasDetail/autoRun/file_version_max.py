import os
import re


def get_version(last_part):
    """从文件名最后一段提取版本号,支持V1和V1(2)格式,返回(主版本, 子版本)"""
    match = re.match(r'V(\d+)(?:\((\d+)\))?', last_part, re.IGNORECASE)
    if match:
        main_ver = int(match.group(1))
        sub_ver = int(match.group(2)) if match.group(2) else 0
        return (main_ver, sub_ver)
    return (0, 0)


def main():
    folder = r'D:\test'
    if not os.path.exists(folder):
        print('文件夹不存在: ' + folder)
        return

    # 获取所有pdf文件
    all_files = [f for f in os.listdir(folder) if f.upper().endswith('.PDF')]

    # 1. 删除文件名中包含SAMPLE的文件
    for f in all_files[:]:
        if 'SAMPLE' in f.upper():
            os.remove(os.path.join(folder, f))
            all_files.remove(f)
            print('已删除(SAMPLE): ' + f)

    # 2. 按第二位数字分组,保留最大版本
    # 文件名格式: PO-{数字}-...-V{版本号}.pdf
    file_groups = {}
    for f in all_files:
        name = os.path.splitext(f)[0]
        parts = name.split('-')
        if len(parts) < 3:
            continue
        # 第二位数字作为分组key
        group_key = parts[1]
        version = get_version(parts[-1])
        if group_key not in file_groups:
            file_groups[group_key] = []
        file_groups[group_key].append((version, f))

    # 每组只保留最大版本,删除其余
    for group_key, file_list in file_groups.items():
        if len(file_list) <= 1:
            continue
        # 按版本号降序排序,第一个为最大版本
        file_list.sort(key=lambda x: x[0], reverse=True)
        for version, f in file_list[1:]:
            os.remove(os.path.join(folder, f))
            print('已删除(版本较低): ' + f)

    print('操作完成!')


if __name__ == '__main__':
    main()
