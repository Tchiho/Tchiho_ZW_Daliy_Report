import matplotlib.pyplot as plt
import pandas as pd
from plottable import ColumnDefinition, ColDef, Table
import numpy
import matplotlib.cm
from datetime import datetime, timedelta
import MySQL

matplotlib.rcParams['font.sans-serif'] = ['SimHei']  # 指定默认字体
matplotlib.rcParams['axes.unicode_minus'] = False  # 解决保存图像时负号'-'显示为方块的问题
Table_1_col_mapping = {"装移-FTTR-昨日归档": "归档", 
                       "装移-FTTR-昨日归档（有连接）": "归档\n（有连接）", 
                       "装移-FTTR-昨日归档（光连接）": "归档\n（光连接）", 
                       "装移-FTTR-今日在途": "FTTR\n在途", 
                       "装移-全量-昨日归档": "全量\n归档",
                       "装移-全量-今日在途": "全量\n在途", 
                       "装移-全量-今日超时": "全量\n超时", 
                       "故障-综维-昨日归档": "综维\n归档", 
                       "故障-综维-今日在途": "综维\n在途", 
                       "故障-综维-今日超时": "综维\n超时", 
                       "故障-装维-昨日归档": "装维\n归档", 
                       "故障-装维-今日在途": "装维\n在途", 
                       "故障-装维-今日超时": "装维\n超时", 
                       "调度-今日在途": "在途", 
                       "调度-今日超时": "超时", 
                       "修障不满意": "不满意"}

def Draw_Table_1():
    data = MySQL.Table_1.values.astype(int)
     # 列的名字
    col_name = ["归档", "归档\n（有连接）", "归档\n（光连接）", "FTTR\n在途", "全量\n归档", "全量\n在途", "全量\n超时", "综维\n归档", "综维\n在途", "综维\n超时", "装维\n归档", "装维\n在途", "装维\n超时", "在途", "超时", "不满意"]
    # 行的名字
    row_name = MySQL.Table_1.index.to_list()
    # 生成一个包含随机数据的表格
    new_df = pd.DataFrame(data, columns=col_name, index=row_name)
    fig, ax = plt.subplots(figsize=(12, 6))
    column = ([ColDef("index", title="分公司", textprops={"ha": "center"}, border='both')]
            +
            [ColumnDefinition(name="归档", group="装移-FTTR", border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="归档\n（有连接）", group='装移-FTTR', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="归档\n（光连接）", group='装移-FTTR', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="FTTR\n在途", group='装移-FTTR', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="全量\n归档", group="装移-全量", border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="全量\n在途", group='装移-全量', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="全量\n超时", group='装移-全量', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="综维\n归档", group='故障', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="综维\n在途", group='故障', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="综维\n超时", group='故障', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="装维\n归档", group="故障", border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="装维\n在途", group='故障', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="装维\n超时", group='故障', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="在途", group="调度", border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="超时", group='调度', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name="不满意", border='both', textprops={"ha": "center"})])
    # 基于pandas表格数据创建和展示图形表格
    Table(new_df, column_definitions=column, row_dividers=True, footer_divider=True)
    # 保存图片
    plt.savefig("./Pic/Table_1.png", dpi=1200, bbox_inches='tight')

def Draw_Table_2():
    time1 = ['0', '0', '0', '0', '0', '0', '0', '0', '0', '0']
    current_date = datetime.now()
    today = current_date.date()
    for i in range(0, 5):
        time1[i] = (today - timedelta(days=i)).strftime("%m-%d")
        time1[i+5] = " " + time1[i] + " "
    time = today - timedelta(days=0)
    data = MySQL.Table_2.values.astype(int)
    # 列的名字
    lie_name = [time1[0], time1[1], time1[2], time1[3], time1[4], time1[5], time1[6], time1[7], time1[8], time1[9]]
    # 行的名字
    hang_name = MySQL.Table_2.index.to_list()
    # 生成一个包含随机数据的表格
    d = pd.DataFrame(data, columns=lie_name, index=hang_name).round(2)
    fig, ax = plt.subplots(figsize=(8.5, 3.5))
    column = ([ColDef("index", title="单位", textprops={"ha": "center"}, border='both')]
            +
            [ColumnDefinition(name=time1[0], group="装移在途超时工单", border='both', textprops={"ha": "center"}),
            ColumnDefinition(name=time1[1], group='装移在途超时工单', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name=time1[2], group='装移在途超时工单', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name=time1[3], group='装移在途超时工单', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name=time1[4], group='装移在途超时工单', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name=time1[5], group="故障在途超时工单", border='both', textprops={"ha": "center"}),
            ColumnDefinition(name=time1[6], group='故障在途超时工单', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name=time1[7], group='故障在途超时工单', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name=time1[8], group='故障在途超时工单', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name=time1[9], group='故障在途超时工单', border='both', textprops={"ha": "center"})])
    # 基于pandas表格数据创建和展示图形表格
    Table(d, column_definitions=column, row_dividers=True, footer_divider=True)

    # 保存图片
    plt.savefig("./Pic/Table_2.png", dpi=1200, bbox_inches='tight')

def Draw_Table_3():
    Tab3 = MySQL.Table_3
    data = Tab3.values.astype(float)
    # 列的名字
    lie_name = ["派单\n数量", "已修复", "未修复", '在途', '感知\n修复', '派单\n数量`', '已修复`', '未修复`', '在途`', '质差\n修复', '总质\n差率', '全市\n在途']
    # 行的名字
    hang_name = MySQL.Table_3.index.to_list()
    # 生成一个包含随机数据的表格
    d = pd.DataFrame(data, columns=lie_name, index=hang_name).round(2)
    fig, ax = plt.subplots(figsize=(10, 4.5))
    column = ([ColDef("index", title="单位", textprops={"ha": "center"}, border='both')]
            +
            [ColumnDefinition(name='派单\n数量', group="感知", border='both', textprops={"ha": "center"}),
            ColumnDefinition(name='已修复', group='感知', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name='未修复', group='感知', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name='在途', group='感知', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name='感知\n修复', group='感知', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name='派单\n数量`', group="质差", border='both', textprops={"ha": "center"}),
            ColumnDefinition(name='已修复`', group='质差', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name='未修复`', group='质差', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name='在途`', group='质差', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name='质差\n修复', group='质差', border='both', textprops={"ha": "center"}),
            ColumnDefinition(name='总质\n差率', group="总整治成功率", border='both', textprops={"ha": "center"}),
            ColumnDefinition(name='全市\n在途', group='总整治成功率', border='both', textprops={"ha": "center"})])

    # 基于pandas表格数据创建和展示图形表格
    tab = Table(d, column_definitions=column, row_dividers=True, footer_divider=True)
    # 保存图片
    plt.savefig("./Pic/Table_3.png", dpi=600, bbox_inches='tight')
