import pandas as pd
from openpyxl import load_workbook
from datetime import date
# 读取Excel文件中的Activity工作表
df = pd.read_excel('myLife.xlsx', sheet_name='Activity')
last_day = df['日期'].iloc[-1]  # 获取最后一天的日期
last_day_operations = df.loc[df['日期'] == last_day, '操作']    # 获取最后一天的操作内容
total_operations = len(last_day_operations) # 统计操作内容的总数
health_score = 0    # 定义健康值变量
# 判断操作总数，并赋予相应的健康值
# 计算步行、喝水、睡眠、吃饭对健康值的贡献，以及类似
# 假设小于1000ml贡献0分，大于2000贡献10分
if total_operations >= 20:
    health_score = 60
else:
    health_score = 60 - 3 * (20 - total_operations)
print(health_score)
# print("最后一天的日期：", last_day)
# print("操作内容的总数：", total_operations)
# print("健康值：", health_score)
 # 获取最新一天操作内容为"喝水"的字段和对应的数量
last_day_water = df.loc[(df['日期'] == last_day) & (df['操作'] == '喝水'), '数量']
last_day_walk = df.loc[(df['日期'] == last_day) & (df['操作'] == '步行'), '数量']
last_day_sleep = df.loc[(df['日期'] == last_day) & (df['操作'].isin(['睡眠', '午睡'])), '时长'].sum()
last_day_meals = df.loc[(df['日期'] == last_day) & (df['操作'].isin(['吃早餐', '吃午餐', '吃晚餐'])), '操作'].count()

print( last_day_meals)
#定义喝水健康值的计算函数
def water_score():
    water_health_score = 0
    if not last_day_water.empty:
        water_quantity = last_day_water.iloc[0]
        if 1000 <= water_quantity <= 2300:
            water_health_score = (water_quantity - 1000) // 130
        elif water_quantity < 1000:
            water_health_score = (water_quantity - 900) // 100 - 1
        elif water_quantity > 2500 and water_quantity <= 4500:
            water_health_score = 10 - (water_quantity - 2500) // 100
        elif water_quantity > 4500:
            water_health_score = -10.0
    return water_health_score
water = water_score()           
print("喝水健康值：", water)
#定义步行健康值的计算函数
def walk_score():
    walk_score = 0
    if not last_day_walk.empty:
        walk_quantity = last_day_walk.iloc[0]
        if 5000 <= walk_quantity <= 10000:
            walk_score = (walk_quantity - 5000) // 625
        elif walk_quantity < 5000:
            walk_score = (walk_quantity - 4500) // 625 - 1
        elif walk_quantity > 10000:
            walk_score = 8.0
    return walk_score
walk = walk_score()           
print("步行健康值：", walk)
#定义睡眠健康值的计算函数
def sleep_score():
    sleep_score = 0
    if last_day_sleep != 0:
        sleep_quantity = last_day_sleep
        if 390 <= sleep_quantity <= 510:
            sleep_score = 12.0
        elif sleep_quantity < 360:
           sleep_score = round((360 - sleep_quantity) / 36, 1)
        elif 510 < sleep_quantity <= 750:
            sleep_score = round(12 - (sleep_quantity - 510) / 24, 1)
        elif sleep_quantity > 750:
            sleep_score = -12.0
    return sleep_score
sleep = sleep_score()           
print("睡眠健康值：", sleep)
#定义吃饭健康值的计算函数
def meals_score():
    meals_score = 0
    if last_day_meals != 0:
        sleep_count = last_day_meals
        if sleep_count == 3:
            meals_score = 10.0
        elif sleep_count == 2:
           meals_score = 8.0
        elif sleep_count == 1:
            meals_score = 0.0
        elif sleep_count == 0:
            meals_score = -12.0
    return meals_score
meals = meals_score()           
print("吃饭健康值：", meals)
health_score =health_score + meals + sleep + walk + water #计算健康值
print("健康值：", health_score)
# 将内容写进指标sheet表
workbook = load_workbook('myLife.xlsx')# 打开Excel文件
worksheet = workbook['指标']# 选择"指标"工作表
last_row = worksheet.max_row + 1# 获取最后一行的行号
# 将日期和健康分数写入Excel文件
worksheet.cell(row=last_row, column=1, value=last_day)
worksheet.cell(row=last_row, column=2, value='健康值')
worksheet.cell(row=last_row, column=3, value=health_score)
# 保存Excel文件
workbook.save('myLife.xlsx')

# 下周计划：