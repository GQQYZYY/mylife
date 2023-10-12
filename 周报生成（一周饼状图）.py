import pandas as pd
from docx import Document
from datetime import datetime, timedelta
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import matplotlib.pyplot as plt
from io import BytesIO
import warnings

# 读取Excel文件
df = pd.read_excel('myLife.xlsx')

# 确保日期列为日期格式
df['日期'] = pd.to_datetime(df['日期'])

# 获取当前日期并筛选前7天的活动数据
current_date = datetime.now()
start_date = current_date - timedelta(days=7)
end_date = current_date
df_7days = df[(df['日期'] >= start_date) & (df['日期'] <= end_date)]

# 使用pivot_table函数按操作和日期进行透视，并计算总时长和出现次数
pivot_df = pd.pivot_table(df_7days, values='时长', index='操作', columns='日期', aggfunc='sum')
pivot_df.columns = pivot_df.columns.strftime('%Y-%m-%d')  # 将日期格式化为字符串

# 计算每个操作在每个日期中的时长所占的百分比
pivot_df_percentage = pivot_df.apply(lambda x: x / x.sum(), axis=0)

# 创建新的Word文档
doc = Document()

# 添加标题
title = doc.add_heading('以下是本周的操作统计', level=1)

# 设置标题居中对齐
title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
# 创建表格
table = doc.add_table(rows=pivot_df_percentage.shape[0]+2, cols=pivot_df_percentage.shape[1]+1)

# 将日期转换为星期几并添加到表头
for i, column in enumerate(pivot_df_percentage.columns):
    date = datetime.strptime(column, '%Y-%m-%d')
    day_of_week = date.strftime('%A')
    table.cell(0, i+1).text = f'{column} ({day_of_week})'

# 将日期转换为星期几并添加到每行数据中
for i, index in enumerate(pivot_df_percentage.index):
    table.cell(i+1, 0).text = index
    for j, column in enumerate(pivot_df_percentage.columns):
        date = datetime.strptime(column, '%Y-%m-%d')
        day_of_week = date.strftime('%A')
        table.cell(i+1, j+1).text = f'{pivot_df_percentage.loc[index, column]:.2%} ({day_of_week})'


# 添加行名和数据
for i, index in enumerate(pivot_df_percentage.index):
    table.cell(i+1, 0).text = index
    for j, column in enumerate(pivot_df_percentage.columns):
        table.cell(i+1, j+1).text = '{:.2%}'.format(pivot_df_percentage.loc[index, column])

# 添加总时长行
table.cell(pivot_df_percentage.shape[0]+1, 0).text = '总时长'
for j, column in enumerate(pivot_df_percentage.columns):
    total_duration = pivot_df[column].sum()
    table.cell(pivot_df_percentage.shape[0]+1, j+1).text = str(total_duration)

# 计算每个操作在7天总时长所占的百分比
total_duration = pivot_df.sum(axis=1)
percentage = total_duration / total_duration.sum()

# 创建饼图
plt.rcParams['font.sans-serif'] = 'SimSun' # SimSun是一个包含中文字符的字体
fig, ax = plt.subplots()  # 修改figsize参数的值来调整图像尺寸
ax.pie(percentage, labels=percentage.index, autopct='%1.1f%%', textprops={'fontsize': 8})
ax.axis('equal')  # 让饼图为圆形

plt.legend(bbox_to_anchor=(1, 1))  # 显示图例
# 保存饼图
plt.savefig('pie_chart.png')
plt.savefig('pie_chart.png', bbox_inches='tight')
# 在Word文档中插入饼图
doc.add_picture('pie_chart.png')

# 格式化文档名称
start_date_str = start_date.strftime('%Y-%m-%d')
end_date_str = end_date.strftime('%Y-%m-%d')

# 保存Word文档
doc.save(f'周报{start_date_str}_{end_date_str}.docx')