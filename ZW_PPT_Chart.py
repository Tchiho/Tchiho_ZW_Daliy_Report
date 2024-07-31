import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import MySQL
def Draw_Chart_1():
    font_path = "./Font/msyh.ttc"  # 微软雅黑字体路径
    chinese_font = FontProperties(fname=font_path, size=8)
    chinese_font1 = FontProperties(fname=font_path, size=10)
    # 准备数据
    categories = MySQL.gz_person
    values = MySQL.gz_num
    plt.figure(figsize=(6, 3))

    # 绘制柱状图
    plt.bar(categories, values)

    # 去掉顶端和右侧的边框
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)

    # 设置标题
    plt.title('个人故障在途工单', fontproperties=chinese_font1)

    # 显示柱状图的高度文本
    for i in range(len(categories)):
        plt.text(categories[i], values[i], values[i], va="bottom", ha="center", fontsize=10)

    # 设置横坐标刻度标签，并使用换行
    # 同时旋转标签以便于阅读
    plt.xticks(categories, fontproperties=chinese_font)
    plt.yticks([0, 5], fontproperties=chinese_font)
    plt.savefig('./Pic/Chart_1.png', dpi=600)

def Draw_Chart_2():
    font_path = "./Font/msyh.ttc"  # 微软雅黑字体路径
    chinese_font = FontProperties(fname=font_path, size=8)
    chinese_font1 = FontProperties(fname=font_path, size=10)
    # 准备数据
    categories = MySQL.zy_person
    values = MySQL.zy_num
    plt.figure(figsize=(4, 3))
    ax = plt.gca()  # 获取当前轴

    # 绘制柱状图
    plt.bar(categories, values, width=0.3)

    # 去掉顶端和右侧的边框
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)

    # 设置标题
    plt.title('个人装移机在途工单', fontproperties=chinese_font1)

    # 显示柱状图的高度文本
    for i in range(len(categories)):
        plt.text(categories[i], values[i], values[i], va="bottom", ha="center", fontsize=10)

    # 设置横坐标刻度标签，并使用换行
    # 同时旋转标签以便于阅读
    plt.xticks(categories, fontproperties=chinese_font)
    plt.yticks([0, 30], fontproperties=chinese_font)
    plt.savefig('./Pic/Chart_2.png', dpi=600)


