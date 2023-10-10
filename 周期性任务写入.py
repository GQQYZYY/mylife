# 本程序将周期性任务（读取Task工作表）分解为各天的具体任务（写到TaskList工作表）
from datetime import datetime, timedelta, date
import pandas as pd  # 导入数据分析模块
from openpyxl import load_workbook  # 写入Excel引擎

df = pd.read_excel("myLife.xlsx", sheet_name="Task")  # 读取Task工作表数据
# df = df[df['次数']>0] # 筛选没有分解过的数据
lastRow = df.shape[0]  # 返回df的行数，类似于len(df); df.shape[1]代表列数

# 读出TaskList中数据，获取现有数据的行数，以便追加数据。
df_tl = pd.read_excel("myLife.xlsx", sheet_name="TaskList")
newRow = df_tl.shape[0]

df_t2 = pd.DataFrame()  # 创建一个空的DataFrame

k = 0  # 临时变量 初始值为0，然后每条课表重复次数累计
for j in range(lastRow):  # 遍历task表的每一行
    df_tl = pd.read_excel("myLife.xlsx", sheet_name="TaskList")
    newRow = df_tl.shape[0]
    df_t2 = pd.DataFrame()
    for i in range(df.iat[j, 6]):  # 遍历任务重复次数
        df_t2.at[i, "日期"] = pd.to_datetime(df.iat[j, 0]) + timedelta(
            days=int(df.iat[j, 5] * i)
        )
        # df_t2.at[i, '日期'] = df.iat[j, 0] + pd.Timedelta(pd.offsets.Day(df.iat[j, 5] * i))
        df_t2.at[i, "开始时间"] = df.iat[j, 1]  # 开始时间
        df_t2.at[i, "操作"] = str(df.iat[j, 2])  # 操作
        df_t2.at[i, "时长"] = df.iat[j, 3]  # 时长
        df_t2.at[i, "对象名"] = df.iat[j, 4]  # 对象名
        df_t2.at[i, "数量"] = ""
        df_t2.at[i, "说明"] = df.iat[j, 7]  # 说明
    k = k + df.iat[j, 6]  # 课表重复次数累计
    df.iat[j, 6] = 0  # 将“Task”中的重复次数清零，以防以后重复往TaskList表里增加记录
    newRow += df.iat[j, 6]

    # 建立写入对象,任务数据的写入
    with pd.ExcelWriter(
        "myLife.xlsx", engine="openpyxl", mode="a", if_sheet_exists="overlay"
    ) as writer:
        df_t2.to_excel(
            writer,
            sheet_name="TaskList",
            startrow=newRow + 1,
            index=False,
            header=False,
        )  # 索引和表头都不写
        df.to_excel(writer, sheet_name="Task", index=False)  # 不写索引
    newRow = df_t2.shape[0]
