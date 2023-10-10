import re
import xlwings as xw
import pandas as pd
import datetime
from openpyxl import load_workbook  # 写入Excel引擎


# 获取课表文件，提取课表数据
def get_course(file_name):
    app = xw.App(visible=True, add_book=False)
    workbooks = app.books.open(file_name)
    # 提取课表数据返回dataframe格式
    data = (
        workbooks.sheets[0]
        .range("A3")
        .options(pd.DataFrame, expand="table", index=True)
        .value
    )
    # 删除课表信息最后一行的备注
    data.drop(data.tail(1).index, inplace=True)
    workbooks.close()
    app.quit()
    return data


# 提取课程名称、上课时间、上课周数信息
def course_information(course):
    # 将一门课程包含的所有信息划分开，存放在列表中
    pat_name = re.compile(r"(.*)\n.*\n.*周")
    name = re.findall(pat_name, course)
    pat_weeks = re.compile(r"(\d.*)周")
    weeks = re.findall(pat_weeks, course)
    pat_times = re.compile(r"[周[](\d\d).*(\d\d)节")
    times = re.findall(pat_times, course)
    return name, weeks, times


# 讨论周数表达的各种情况，返回格式如[('7','14'),'18']
def format_weeks(week):
    if "," in week and "-" not in week:
        n = re.split(",", week)
    if "-" in week and "," not in week:
        pat = re.compile(r"(\d*)-(\d*)")
        n = re.findall(pat, week)
    if "," not in week and "-" not in week:
        n = [week]
    if "," in week and "-" in week:
        week = week.split(",")
        n = []
        for j in week:
            if "-" in j:
                pat = re.compile(r"(\d*)-(\d*)")
                m = re.findall(pat, j)
                n.extend(m)
            else:
                n.append(j)
    return n


# 将提取的周数转为起始终止日期
def weeks_to_date(init_time, week):
    # 设置初始上课时间，第一周的星期一日期为2023-9-4
    date_ = []
    for m in week:
        if type(m) is str:
            start_time = init_time + datetime.timedelta(days=7 * (int(m) - 1))
            end_time = start_time + datetime.timedelta(days=6)
            date_.append([start_time, end_time])
        elif type(m) is tuple:
            start_time = init_time + datetime.timedelta(days=7 * (int(m[0]) - 1))
            end_time = init_time + datetime.timedelta(days=7 * int(m[1]))
            end_time = end_time + datetime.timedelta(days=-1)
            date_.append([start_time, end_time])
    return date_


# 确定上课时间与计算上课时长
def calculate_time(times):
    if times[0] == "01":
        begin_time = "8:20"
    elif times[0] == "03":
        begin_time = "10:15"
    elif times[0] == "05":
        begin_time = "14:30"
    elif times[0] == "07":
        begin_time = "16:25"
    else:
        begin_time = "19:00"
    time = (int(times[1]) - int(times[0]) + 1) * 45
    return begin_time, time


# 新建dataframe，添加课程信息，并删除冗余课程
def preprocessing_course(data, init_date):
    df = pd.DataFrame(columns=["日期", "开始时间", "操作", "时长", "对象名", "说明"])
    day = 0
    # items()方法用于迭代DataFrame中的每一列，它返回一个包含（列名，列数据）的元组迭代器。
    for col in data.items():
        print("+++++++++++++++++++++")
        print(col)
        for i in range(len(col[1])):
            print("------------------------------")
            if col[1][i] is None:
                continue
            names, weeks_, times = course_information(col[1][i])
            for j in range(len(names)):
                print("********************************")
                begin_, course_time = calculate_time(times[j])
                week = format_weeks(weeks_[j])
                print(week)
                date_ = weeks_to_date(init_date, week)
                print(date_)
                for lis in date_:
                    start_date = lis[0] + datetime.timedelta(days=day)
                    delta = datetime.timedelta(days=7)
                    date = start_date
                    while date <= lis[1]:
                        df = df._append(
                            pd.Series(
                                [date, begin_, "上课", course_time, names[j], ""],
                                index=["日期", "开始时间", "操作", "时长", "对象名", "说明"],
                            ),
                            ignore_index=True,
                        )
                        date += delta
        day += 1
    df = df.drop_duplicates(subset=["日期", "开始时间"], keep="first")
    df = df.sort_values(by=["日期"])
    return df


# 写入TaskList
def write_course(data, filename):
    df_tl = pd.read_excel("myLife.xlsx", sheet_name="TaskList")
    newRow = df_tl.shape[0]

    with pd.ExcelWriter(
        "myLife.xlsx", engine="openpyxl", mode="a", if_sheet_exists="overlay"
    ) as writer:
        df.to_excel(
            writer,
            sheet_name="TaskList",
            startrow=newRow + 1,
            index=False,
            header=False,
        )  # 索引和表头都不写


if __name__ == "__main__":
    data = get_course("23软件工程1班.xls")
    init_date = datetime.datetime(2023, 9, 4)
    df = preprocessing_course(data, init_date)
    write_course(df, "myLife.xlsx")
