# -*- coding:utf-8 -*-
import random

class VAS_GUI():
    # 获取随机名单
    def get_files(self, strings, group_size):
        random.shuffle(strings)
        grouped_strings = [strings[i:i+group_size] for i in range(0, len(strings), group_size)]
        return grouped_strings


def gui_start():
    VAS = VAS_GUI()
    input_strings = ["宋云生","王玉国","李安平","彭俊尧","刘倩","宁宇杰","张晓君","冮宁","张素梅","田佳桐","杨娣","季毅","孙适","戴世强","龚建杉","车帅","张岩","黄雪","邵秀清","王晓庆","徐薇","梅新鑫","刘晓凤","张佳雯","李欣阳","仪德宇","杨丹","郭晓雨","汪明珠","杜雨燕","林晓光","姜亚维","吴朔","孙洪义","曲晓璞","李娟","张晶","丛宇婷","张天媛","孟琳","杨雪","任晓晨","林彤彤","张昕","单晓","王贺","田芳","刘凤俏","张巍旭","李敏","吴晓丹","杨雪","宋清云","刘宏","王楠","王黎明","郭欣宜","纪红梅","田源","薛敏","沈静","刘金彤","刘彬","吕明慧","崔小丽","韩冲","杜双","李佳音","李士闯","张桂芹","邹越","郑慧","李晶","邵立影","张美双","于苓","张宇","葛明晶","张千里","高明","纪元主","赵文刚","梁伟光","金志海","巴春凤","吕俊","姜燕","石元峰","吴立强","刘新","公兵","矫世军","史运刚","赖文光","宋宝龙","高宪鹏","张云彤","潘恩磊","孙兴文","于英杰","范庆春","杨以梅","刘艳","解华丽","宋岩","尹士霞","王玉珍","汤伟","冯金丽","谷美红","朱叶","周晶晶","张汀汀","刘姣","孙华","赵青","宋进举","张延平","吴一笛","刘顺发","周福禄","李硕","吴珊珊","汪玉","于秀红","李海丹","初文杰","杨秀坤","钟吉莅","赵程程","王有喜","宋心宇","高妮","王倩倩","宫殿嫔","邓晓娜","高林艳","刘欣妍","赵晶晶","乔元鑫","赵春艳","肖娜","王威","马俊威","赵锋","刘亚辉","肖阿婷","刘可心","陈凤洁","汪雨潼","于晓慧","姚钧译","陈明昊","赖总","姜姐"]

    group_size = 10
    result = VAS.get_files(input_strings, group_size)

    # 处理剩余的不足10个字符串
    if len(result[-1]) < group_size:
        remaining_strings = result.pop()
        while remaining_strings:
            group = remaining_strings[:group_size]
            result.append(group)
            remaining_strings = remaining_strings[group_size:]

    for group in result:
        print(group)


if __name__ == '__main__':
    gui_start()
