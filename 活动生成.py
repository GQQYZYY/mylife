# 从TaskList中获取某日的任务，并通过Activity历史活动数据，预测该日较大概率发生的活动，不覆盖写入Activity工作表
import datetime  # 导入获取今日日期的内置模块
import pandas as pd  # 导入数据分析第三方库
from openpyxl import load_workbook  # 导入访问Excel的模块openpyxl
import random

# 确定日期
thisDate = datetime.date.today() - datetime.timedelta(
    days=0
)  # 前几天的日期，今天、昨天、前天分别取days=0、1、2
strDate = thisDate.strftime("%Y-%m-%d")  # 转换为字符串
print(thisDate)

df = pd.read_excel("myLife.xlsx", sheet_name="Activity")  # 获取Activity历史数据
lastRow = df.shape[0]  # Activity工作表的行数，df.shape[1]为列数

# 这一段代码从TaskList中获取某日的任务
df_task = pd.read_excel("myLife.xlsx", sheet_name="TaskList")  # 获取TaskList任务
df_task = df_task.loc[df_task["日期"] == strDate]  # 查询出该日的任务
df_task["日期"] = thisDate  # 保留日期，不要后面的几点几分
print(df_task)
taskRow = df_task.shape[0]  # task行数
print(taskRow)

# 这一段代码通过Activity历史活动数据，预测该日较大概率发生的活动
df = df.groupby("操作", as_index=False).agg(
    {"日期": "count", "时长": "mean", "数量": "mean"}
)  # 对操作分组统计
# df['时长'] = df['时长'].astype(int) # 把‘时长’这列转化为整数型
df["时长"] = df["时长"].fillna(0).astype(int)  # 对非空数据取整
df["数量"] = (
    df["数量"].fillna(0).astype(int).astype(object).where(df["数量"].notnull())
)  # 对非空数据取整
df.drop(df[df["操作"] == "上课"].index, inplace=True)  # 不考虑上课活动
df.sort_values("日期", ascending=False, inplace=True)  # 对‘日期’统计的次数进行降序排序
df = df.head(29)  # 取前N行
# df.loc[df['操作']=='步行']['数量']=random.randint(6000,10000)
# df['数量'].where(df['操作'] != '步行',other=random.randint(6000,10000), inplace=True)
print(df)
# 下面的转换使得df与Activity表的各列对应上
df = df.drop(columns="日期")  # 删除“日期”这列
df.insert(0, "日期", thisDate)  # 在位置0插入列，设置列名为“日期”，此列内容为todays_date
df.insert(1, "开始时间", "")  # 在位置1插入列，设置列名为“开始时间”，此列内容为空，用‘’表示
df.insert(4, "对象", "")  # 插入空列
df.insert(6, "说明", "")  # 插入空列

# 建立写入对象，写入活动数据
with pd.ExcelWriter(
    "myLife.xlsx", engine="openpyxl", mode="a", if_sheet_exists="overlay"
) as writer:
    df_task.to_excel(
        writer, sheet_name="Activity", startrow=lastRow + 1, index=False, header=False
    )  # 将数据写入excel中的aa表,从第一个空行开始写
    df.to_excel(
        writer,
        sheet_name="Activity",
        startrow=lastRow + taskRow + 1,
        index=False,
        header=False,
    )  # 将数据写入excel中的Activity表,从第一个空行开始写
