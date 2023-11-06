import openpyxl
from datetime import datetime, timedelta

# 打开 Excel 文件
workbook = openpyxl.load_workbook('myLife.xlsx')  # 替换为您的 Excel 文件名或路径

# 选择工作表
sheet = workbook['计划']  # 替换为您的工作表名称
now_date = datetime.now()  # 获取当前日期时间
new_date = now_date - timedelta(days=1)  # 获取昨天的日期
new_date = new_date.strftime('%Y-%m-%d')  # 将日期转换为字符串格式
now_date = now_date.strftime('%Y-%m-%d')  # 将日期时间转换为字符串格式
print(now_date)  # 打印当前日期

# 获取表格中的数据
data = []
for row in sheet.iter_rows(values_only=True):
    data.append(row)
# 查找并更新数据
data1 = {}  # 用于追踪已经出现的计划名称
for row in data[1:]:
    plan_type = row[0]
    if plan_type == '短期计划':
        name = row[1]  # 获取计划名称
        if name in data1:
            # 如果计划名称已经存在于字典中，更新字典的值为当前行的数据
            data1[name] = row
        else:
            # 如果计划名称不存在于字典中，将当前行的数据添加到字典中
            data1[name] = row
data1 = list(data1.values())
print(type(data1))

# 查找并计算百分比
for row in data1:
    plan_type = row[0]  # 假设“计划类型”在第一列
    name = row[1]
    current_date = row[4]  # 获取当前日期
    completion = row[8]  # 获取完成度
    start_date = row[3]  # 获取开始日期
    end_date = row[5]  # 获取结束日期
    finish = row[7]
    finish = float(finish)  # 将完成度转换为浮点数
    # 输出完成百分比
    print("finish", finish)
    plan = row[6]
    plan = int(plan)  # 将计划转换为整数
    row = list(row)  # 将行数据转换为列表
    duration = (datetime.strptime(end_date, '%Y-%m-%d') - datetime.strptime(start_date, '%Y-%m-%d')).days + 1
    print("duration", duration)
    one_day = round((100 / duration), 1)
    print("one_day", one_day)
    # 解析字符串计划和持续时间，并将其转换为整数
    # 计算完成百分比并四舍五入到一位小数
    finish = round((finish + plan / float(duration)), 1)
    print("finish", finish)
    if  current_date < now_date <= end_date :
        completion = round((float(completion) + one_day),1)
        print("completion", completion)
    # 更新“完成度”的值
        row[8] = str(completion)  # 将更新后的完成度值转换为字符串并赋值给对应单元格
        # 更新“当日日期”的值
        row[4] = now_date  # 将当前日期赋值给当日日期单元格
        row[7] = finish
        print(row)
        sheet.append(row)  # 将行数据添加到表格的最后一行

# 保存修改后的 Excel 文件
workbook.save('myLife.xlsx')  # 替换为您的 Excel 文件名或路径