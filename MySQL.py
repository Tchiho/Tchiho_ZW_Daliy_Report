import pymysql
import pandas
from datetime import datetime, timedelta

# 设置MySQL连接信息
MYSQL_USER = 'report'
MYSQL_PASSWORD = 'Ev9@EPQ_Zt'
MYSQL_HOST = '133.0.132.51'
MYSQL_PORT = 54588
MYSQL_DB = 'report'

# 获取当前日期
current_date = datetime.now()
# 获取中文星期的名称
weekday_cn = {0: '星期一', 1: '星期二', 2: '星期三', 3: '星期四', 4: '星期五', 5: '星期六', 6: '星期日'}
weekday_name = weekday_cn[current_date.weekday()]
# 格式化日期为"YYYY年M月D日 星期N"格式
formatted_date = f"{current_date.year}年{current_date.month}月{current_date.day}日 {weekday_name}"
today = current_date.date()


# 连接到 MySQL 数据库
dbconn = pymysql.connect(host=MYSQL_HOST, port=MYSQL_PORT, user=MYSQL_USER, password=MYSQL_PASSWORD, database=MYSQL_DB)
# 建立游标
cursor = dbconn.cursor()
# 创建Pandas空Dataframe， 表1
Table_1 = pandas.DataFrame()
Table_2 = pandas.DataFrame()
Table_3 = pandas.DataFrame()
Txt_2 = ""
Txt_6 = ""
Txt_7 = ""
Units_129 = pandas.read_csv("./Txt/129分支局.txt", delimiter="\t")
# 创建分公司列表
regions = ['孝南', '云梦', '大悟', '新城', '孝昌', '安陆', '汉川', '应城']
# 双表数组
gz_person = []
gz_num = []
zy_person = []
zy_num = []
# 设置表格列
Table_1_new_col = ["装移-FTTR-昨日归档", "装移-FTTR-昨日归档（有连接）", "装移-FTTR-昨日归档（光连接）", "装移-FTTR-今日在途", "装移-全量-昨日归档","装移-全量-今日在途", "装移-全量-今日超时", "故障-综维-昨日归档", "故障-综维-今日在途", "故障-综维-今日超时", "故障-装维-昨日归档", "故障-装维-今日在途", "故障-装维-今日超时", "调度-今日在途", "调度-今日超时", "修障不满意"]
Table_3_new_col = ["感知-派单数量","感知-已修复", "感知-未修复", "感知-在途", "感知-感知修复", "质差-派单数量", "质差-已修复", "质差-未修复", "质差-在途", "质差-感知修复", "总整治成功率-总质差率", "总整治成功率-全市在途"]
# 测试区域



# 执行查询：装移-FTTR-昨日归档……（表68）表1
def select_table68():
    for region in regions:
        index = regions.index(region)
        sql = "SELECT * FROM table68 WHERE col2 LIKE %s AND col3 = '小计' AND col = %s"
        if region == '孝南':
            region = '孝感'
        cursor.execute(sql, (region + '%', today))
        if region == '孝感':
            region = '孝南'
        # 获取查询结果
        results = cursor.fetchall() 
        # 获取元组长度
        result_len = results.__len__()
        if result_len != 1 :
            print("数据库信息错误，检查数据是否正确！")
            Table_1.loc[region, '装移-FTTR-昨日归档'] = 0
            Table_1.loc[region, '装移-FTTR-昨日归档（有连接）'] = 0
            Table_1.loc[region, '装移-FTTR-昨日归档（光连接）'] = 0

        # 打印结果
        count_col1 = 0
        count_col2 = 0
        count_col3 = 0
        for row in results:
            count_col1 = row[12] + count_col1
            count_col2 = row[13] + row[14] + row[15] + count_col2
            count_col3 = row[14] + count_col3
            Table_1.loc[region, '装移-FTTR-昨日归档'] = count_col1
            Table_1.loc[region, '装移-FTTR-昨日归档（有连接）'] = count_col2
            Table_1.loc[region, '装移-FTTR-昨日归档（光连接）'] = count_col3

# 执行查询：装移-FTTR-今日在途……（装移机在途清单）表1
def select_tablezjlive():
    for region in regions:
        # FTTR今日在途
        sql = "SELECT col1 FROM tablezjlive WHERE col12 LIKE %s AND (col36 LIKE %s OR col36 LIKE %s) AND col = %s"
        cursor.execute(sql, (region + '%', '%全光%', '%FTTR%', today))
        results = cursor.fetchall()
        result_len = results.__len__()
        Table_1.loc[region, '装移-FTTR-今日在途'] = result_len
        # 全量今日在途
        sql = "SELECT col1 FROM tablezjlive WHERE col12 LIKE %s AND col = %s"
        cursor.execute(sql, (region + '%', today))
        results = cursor.fetchall()
        result_len = results.__len__()
        Table_1.loc[region, '装移-全量-今日在途'] = result_len
        # 全量今日超时
        sql = "SELECT col1 FROM tablezjlive WHERE col12 LIKE %s AND col42 = '是' AND col = %s"
        cursor.execute(sql, (region + '%', today))
        results = cursor.fetchall()
        result_len = results.__len__()
        Table_1.loc[region, '装移-全量-今日超时'] = result_len

# 执行查询：装移-全量-昨日归档（装移机归档清单）表1
def select_table3():
    for region in regions:
        sql = "SELECT * FROM table3 WHERE col1 LIKE %s AND col = %s"
        if region == '孝南':
            region = '孝感'
        num = cursor.execute(sql, (region + '%', today))
        if region == '孝感':
            region = '孝南'
        results = cursor.fetchall()
        Table_1.loc[region, '装移-全量-昨日归档'] = num

#执行查询：故障-综维-昨日归档（故障归档清单）表1
def select_table6():
    for region in regions:
        sql = "SELECT * FROM table6 WHERE col2 LIKE %s AND col = %s"
        cursor.execute(sql, ("孝感" + region + "%", today))
        results = cursor.fetchall()
        Table_1.loc[region, '故障-综维-昨日归档'] = results.__len__() 
        sql = "SELECT * FROM table6 WHERE col2 LIKE %s AND col = %s"
        cursor.execute(sql, (region + "%", today))
        results = cursor.fetchall()
        Table_1.loc[region, '故障-装维-昨日归档'] = results.__len__() 

#执行查询：故障-综维-今日在途（故障在途清单）表1
def select_tablegzlive():
    for region in regions:
        sql = "SELECT * FROM tablegzlive WHERE col4 LIKE %s AND col = %s"
        cursor.execute(sql, ("孝感" + region + "%", today))
        results = cursor.fetchall()
        Table_1.loc[region, '故障-综维-今日在途'] = results.__len__()

        sql = "SELECT * FROM tablegzlive WHERE col4 LIKE %s AND col2 = '是' AND col = %s"
        cursor.execute(sql, ("孝感" + region + "%", today))
        results = cursor.fetchall()
        Table_1.loc[region, '故障-综维-今日超时'] = results.__len__() 

        sql = "SELECT * FROM tablegzlive WHERE col4 LIKE %s AND col = %s"
        cursor.execute(sql, (region + "%", today))
        results = cursor.fetchall()
        Table_1.loc[region, '故障-装维-今日在途'] = results.__len__() 

        sql = "SELECT * FROM tablegzlive WHERE col4 LIKE %s AND col2 = '是' AND col = %s"
        cursor.execute(sql, (region + "%", today))
        results = cursor.fetchall()
        Table_1.loc[region, '故障-装维-今日超时'] = results.__len__()
    return 0

#执行查询：调度（调度工单清单）表1
def select_tabledd():
    for region in regions:
        sql = "SELECT * FROM tabledd WHERE col3 LIKE %s AND col = %s"
        cursor.execute(sql, (region + "%", today))
        results = cursor.fetchall()
        result_len = results.__len__()
        Table_1.loc[region, '调度-今日在途'] = result_len

        sql = "SELECT * FROM tabledd WHERE col3 LIKE %s AND col2 LIKE %s AND col = %s"
        cursor.execute(sql, (region + "%", "超%", today))
        results = cursor.fetchall()
        result_len = results.__len__()
        Table_1.loc[region, '调度-今日超时'] = result_len

#执行查询：修障不满意（修障不满意清单）表1
def select_tablexz():
    for region in regions:
        sql = "SELECT * FROM tablexz WHERE col1 LIKE %s AND col = %s"
        num = cursor.execute(sql, (region + "%", today))
        Table_1.loc[region, '修障不满意'] = num

#执行查询：装移机履约率（表1），文本7
def select_table1(txt):
    txt = "装移机履约率指标>=99.8%\n装移机履约率未达标单位:\n"
    state = True
    indexs = Units_129.values.tolist()
    Temp_index = 0.998
    for index in indexs:
        # index = str(index)
        # index = index.replace('[', '').replace(']', '').replace("'", '')
        sql = "SELECT col5,col7 FROM table1 WHERE col5 = %s AND col6 = '小计' AND col7 <= %s AND col = %s"
        cursor.execute(sql, (index[0], Temp_index, today))
        results = cursor.fetchall()
        for result in results:
            state = False
            txt = txt + result[0] + "\t履约率为" + str(result[1]) + "\n"
    if state:
        txt = txt + "无"
    return txt

# 执行查询：感知工单（满意度修复工单详情统计），表3
def select_tablemyd():
    for region in regions:
        sql = "SELECT * FROM tablemyd WHERE col42 LIKE %s AND col = %s"
        num = cursor.execute(sql, (region + "%", today))
        Table_3.loc[region, "感知-派单数量"] = num

        sql = "SELECT * FROM tablemyd WHERE col42 LIKE %s AND col3 = '已修复' AND col = %s"
        num = cursor.execute(sql, (region + "%", today))
        Table_3.loc[region, "感知-已修复"] = num

        sql = "SELECT * FROM tablemyd WHERE col42 LIKE %s AND col40 = 'nan' AND col = %s"
        num = cursor.execute(sql, (region + "%", today))
        Table_3.loc[region, "感知-在途"] = num


        sql = "SELECT * FROM tablemyd WHERE col42 LIKE %s AND col3 = '未修复' AND col = %s"
        num = cursor.execute(sql, (region + "%", today)) - num
        Table_3.loc[region, "感知-未修复"] = num

    Table_3["感知-感知修复"] =Table_3["感知-已修复"] / Table_3["感知-派单数量"]

# 执行查询：质差工单（派单详情），表3
def select_tablezc():
    sql = "DELETE FROM tablezc WHERE col3 = '未修复' AND col24 = '网关承载不足派单'"
    deletenum = cursor.execute(sql)
    for region in regions:
        sql = "SELECT col1 FROM tablezc WHERE col42 LIKE %s AND col = %s"
        num = cursor.execute(sql, (region + "%", today))
        Table_3.loc[region, "质差-派单数量"] = num

        sql = "SELECT col1 FROM tablezc WHERE col42 LIKE %s AND col3 = '已修复' AND col = %s"
        num = cursor.execute(sql, (region + "%", today))
        Table_3.loc[region, "质差-已修复"] = num

        sql = "SELECT col1 FROM tablezc WHERE col42 LIKE %s AND col40 = 'nan' AND col3 = '未修复' AND col = %s"
        num = cursor.execute(sql, (region + "%", today))
        Table_3.loc[region, "质差-在途"] = num

        sql = "SELECT col1 FROM tablezc WHERE col42 LIKE %s AND col3 = '未修复' AND col = %s"
        num = cursor.execute(sql, (region + "%", today)) - num
        Table_3.loc[region, "质差-未修复"] = num

# 文本实现：文本6， 表3
def chart_zc(Table_3):
    Table_3["质差-感知修复"] =Table_3["质差-已修复"] / Table_3["质差-派单数量"]
    Table_3["总整治成功率-全市在途"] = Table_3["感知-在途"] + Table_3["质差-在途"]
    Table_3["总整治成功率-总质差率"] = (Table_3["感知-已修复"] + Table_3["质差-已修复"]) / (Table_3["感知-派单数量"] + Table_3["质差-派单数量"]) 
    sorted_Table_3 = Table_3.sort_values(by = "总整治成功率-总质差率", ascending = True)
    Txt = "{one}、{two}和{three}总质差工单整治成功率排全市后三位".format(one = sorted_Table_3.index[0], two = sorted_Table_3.index[1], three = sorted_Table_3.index[2])
    return Txt

# 文本实现：文本框2
def write_Txt_2(Table_1):
    Table_1 = Table_1.astype(int)
    sorted_Table_1 = Table_1.sort_values(by = "装移-全量-今日在途", ascending = False)
    # 屎山
    col_num_index = sorted_Table_1.columns.get_loc("装移-全量-今日在途")
    col_overtime_index = sorted_Table_1.columns.get_loc("装移-全量-今日超时")
    ztzy_region_1 = sorted_Table_1["装移-全量-今日在途"].index[0]
    ztzy_num_1 = sorted_Table_1.iloc[0, col_num_index]
    ztzy_overtime_1 = sorted_Table_1.iloc[0, col_overtime_index]

    ztzy_region_2 = sorted_Table_1["装移-全量-今日在途"].index[1]
    ztzy_num_2 = sorted_Table_1.iloc[1, col_num_index]
    ztzy_overtime_2 = sorted_Table_1.iloc[1, col_overtime_index]

    ztzy_region_3 = sorted_Table_1["装移-全量-今日在途"].index[2]
    ztzy_num_3 = sorted_Table_1.iloc[2, col_num_index]
    ztzy_overtime_3 = sorted_Table_1.iloc[2, col_overtime_index]
    
    sorted_Table_1["故障-今日在途"] = Table_1["故障-综维-今日在途"] + Table_1["故障-装维-今日在途"]
    sorted_Table_1["故障-今日超时"] = Table_1["故障-综维-今日超时"] + Table_1["故障-装维-今日超时"]
    sorted_Table_1 = sorted_Table_1.sort_values(by = "故障-今日在途", ascending = False)

    col_num_index = sorted_Table_1.columns.get_loc("故障-今日在途")
    col_overtime_index = sorted_Table_1.columns.get_loc("故障-今日超时")

    ztgz_region_1 = sorted_Table_1["故障-今日在途"].index[0]
    ztgz_num_1 = sorted_Table_1.iloc[0, col_num_index]
    ztgz_overtime_1 = sorted_Table_1.iloc[0, col_overtime_index]

    ztgz_region_2 = sorted_Table_1["故障-今日在途"].index[1]
    ztgz_num_2 = sorted_Table_1.iloc[1, col_num_index]
    ztgz_overtime_2 = sorted_Table_1.iloc[1, col_overtime_index]

    ztgz_region_3 = sorted_Table_1["故障-今日在途"].index[2]
    ztgz_num_3 = sorted_Table_1.iloc[2, col_num_index]
    ztgz_overtime_3 = sorted_Table_1.iloc[2, col_overtime_index]

    sorted_Table_1 = Table_1.sort_values(by = "调度-今日在途", ascending = False)

    col_num_index = sorted_Table_1.columns.get_loc("调度-今日在途")
    col_overtime_index = sorted_Table_1.columns.get_loc("调度-今日超时")

    ztdd_region_1 = sorted_Table_1["调度-今日在途"].index[0]
    ztdd_num_1 = sorted_Table_1.iloc[0, col_num_index]
    ztdd_overtime_1 = sorted_Table_1.iloc[0, col_overtime_index]

    ztdd_region_2 = sorted_Table_1["调度-今日在途"].index[1]
    ztdd_num_2 = sorted_Table_1.iloc[1, col_num_index]
    ztdd_overtime_2 = sorted_Table_1.iloc[1, col_overtime_index]

    ztdd_region_3 = sorted_Table_1["调度-今日在途"].index[2]
    ztdd_num_3 = sorted_Table_1.iloc[2, col_num_index]
    ztdd_overtime_3 = sorted_Table_1.iloc[2, col_overtime_index]

    all = Table_1["装移-全量-今日在途"].sum() + Table_1["故障-综维-今日在途"].sum() + Table_1["故障-装维-今日在途"].sum()
    ztzy_all = Table_1["装移-全量-今日在途"].sum()
    ztzy_overtime = Table_1["装移-全量-今日超时"].sum()
    ztgz_all = Table_1["故障-综维-今日在途"].sum() + Table_1["故障-装维-今日在途"].sum()
    ztgz_overtime = Table_1["故障-综维-今日超时"].sum() + Table_1["故障-装维-今日超时"].sum()
    ztdd_all = Table_1["调度-今日在途"].sum()
    ztdd_overtime = Table_1["调度-今日超时"].sum()
    bmy_all = Table_1["修障不满意"].sum()
    Temp_txt = "每日工作温馨提醒:{date}上午9点全市在途装移修全量工单{all}张。\n"\
        "全市在途装移机工单{ztzy_all}张，超时{ztzy_overtime}张；\n"\
        "TOP3:  {ztzy_region_1} {ztzy_num_1}张，超时{ztzy_overtime_1}张；\n"\
        "       {ztzy_region_2} {ztzy_num_2}张，超时{ztzy_overtime_2}张；\n"\
        "       {ztzy_region_3} {ztzy_num_3}张，超时{ztzy_overtime_3}张；\n"\
        "全市在途故障工单{ztgz_all}张，超时{ztgz_overtime}张\n"\
        "TOP3:  {ztgz_region_1} {ztgz_num_1}张， 超时{ztgz_overtime_1}张；\n"\
        "       {ztgz_region_2} {ztgz_num_2}张， 超时{ztgz_overtime_2}张；\n"\
        "       {ztgz_region_3} {ztgz_num_3}张 ，超时{ztgz_overtime_3}张；\n"\
        "全市在途调度工单{ztdd_all}张，超时{ztdd_overtime}张；\n"\
        "TOP3:  {ztdd_region_1} {ztdd_num_1}张， 超时{ztdd_overtime_1}张；\n"\
        "       {ztdd_region_2} {ztdd_num_2}张， 超时{ztdd_overtime_2}张；\n"\
        "       {ztdd_region_3} {ztdd_num_3}张， 超时{ztdd_overtime_3}张；\n"\
        "全市修障不满意需上门修复工单{bmy_all}张；".format(date = formatted_date, all = all, 
                                           ztzy_all = ztzy_all, ztzy_overtime = ztzy_overtime, 
                                           ztzy_region_1 = ztzy_region_1, ztzy_num_1 = ztzy_num_1, ztzy_overtime_1 = ztzy_overtime_1,
                                           ztzy_region_2 = ztzy_region_2, ztzy_num_2 = ztzy_num_2, ztzy_overtime_2 = ztzy_overtime_2,
                                           ztzy_region_3 = ztzy_region_3, ztzy_num_3 = ztzy_num_3, ztzy_overtime_3 = ztzy_overtime_3,
                                           ztgz_all = ztgz_all, ztgz_overtime = ztgz_overtime, 
                                           ztgz_region_1 = ztgz_region_1, ztgz_num_1 = ztgz_num_1, ztgz_overtime_1 = ztgz_overtime_1,
                                           ztgz_region_2 = ztgz_region_2, ztgz_num_2 = ztgz_num_2, ztgz_overtime_2 = ztgz_overtime_2,
                                           ztgz_region_3 = ztgz_region_3, ztgz_num_3 = ztgz_num_3, ztgz_overtime_3 = ztgz_overtime_3,                                           
                                           ztdd_all = ztdd_all, ztdd_overtime = ztdd_overtime, 
                                           ztdd_region_1 = ztdd_region_1, ztdd_num_1 = ztdd_num_1, ztdd_overtime_1 = ztdd_overtime_1,
                                           ztdd_region_2 = ztdd_region_2, ztdd_num_2 = ztdd_num_2, ztdd_overtime_2 = ztdd_overtime_2,
                                           ztdd_region_3 = ztdd_region_3, ztdd_num_3 = ztdd_num_3, ztdd_overtime_3 = ztdd_overtime_3,
                                           bmy_all = bmy_all
                                           )
    return Temp_txt

# 图表实现：个人在途工单人员
def select_chart3():

    sql = "SELECT col3, col4, COUNT(col1) AS count FROM tablegzlive WHERE col = %s GROUP BY col3, col4 ORDER BY count DESC LIMIT 4"
    cursor.execute(sql, today)
    results = cursor.fetchall()
    for result in results:
        gz_person.append(result[0])
        gz_num.append(result[2])

    sql = "SELECT col12, col13, COUNT(col1) AS count FROM tablezjlive WHERE col = %s GROUP BY col12, col13 ORDER BY count DESC LIMIT 4"
    cursor.execute(sql, today)
    results = cursor.fetchall()
    for result in results:
        zy_person.append(result[1])
        zy_num.append(result[2])

# 表格实现：全市在途超时，表2
def select_chart4():
    for index in range(0,5):
        for region in regions:
            sql = "SELECT col1 FROM tablezjlive WHERE col12 LIKE %s AND col42 = '是' AND col = %s"
            time = today-timedelta(days=index)
            num = cursor.execute(sql, (region + '%', time))
            Table_2.loc[region, "装移在途超时工单" + time.strftime("%Y-%m-%d")] = num 
            # Table_2.loc[region, "装" + str(5-index)] = num

    for index in range(0,5):
        for region in regions:
            sql = "SELECT col1 FROM tablegzlive WHERE col4 LIKE %s AND col2 = '是' AND col = %s"
            time = today-timedelta(days=index)
            num = cursor.execute(sql, ("%" + region + '%', time))
            Table_2.loc[region, "故障在途超时工单" + time.strftime("%Y-%m-%d")] = num 
            # Table_2.loc[region, "故" + str(5-index)] = num



# 函数使用区域
# 装维在途工单统计（Table_1）处理区域
select_table68()
select_tablezjlive()
select_table3()
select_table6()
select_tablegzlive()
select_tabledd()
select_tablexz()
select_chart3()
select_chart4()
select_tablemyd()
select_tablezc()

Txt_7 = select_table1(Txt_7)
Txt_6 = chart_zc(Table_3)
Txt_2 = write_Txt_2(Table_1)

def Trans_Table(table, table_new_col):
    for col in table.columns.tolist():
        table.loc["全市", col] = table.loc[:, col].sum()
    table = table[table_new_col]
    return table

def Trans_Table_mini(table):
    for col in table.columns.tolist():
        table.loc["全市", col] = table.loc[:, col].sum()
    return table

# print(Table_1)

Table_1 = Trans_Table(Table_1, Table_1_new_col)
Table_2 = Trans_Table_mini(Table_2)
Table_3 = Trans_Table(Table_3, Table_3_new_col)


Table_3.loc["全市", "感知-感知修复"] =Table_3.loc["全市", "感知-已修复"] / Table_3.loc["全市", "感知-派单数量"]
Table_3.loc["全市", "质差-感知修复"] =Table_3.loc["全市", "质差-已修复"] / Table_3.loc["全市", "质差-派单数量"]
Table_3.loc["全市", "总整治成功率-全市在途"] = Table_3.loc["全市", "感知-在途"] + Table_3.loc["全市", "质差-在途"]
Table_3.loc["全市", "总整治成功率-总质差率"] = (Table_3.loc["全市", "感知-已修复"] + Table_3.loc["全市", "质差-已修复"]) / (Table_3.loc["全市", "感知-派单数量"] + Table_3.loc["全市", "质差-派单数量"]) 



# Table数据类型全部转化为int
Table_1 = Table_1.astype(int)
Table_2 = Table_2.astype(int)
# 保留两位小数
Table_3 = Table_3.round(2)

# print(Table_2)
cursor.close()
dbconn.close()