# 本程序打开当前文件夹中的myLife.xlsx文件，在Scene工作表中写入今天的日期、本人所在城市、天气信息等。
import cpca  # 导入chinese_province_city_area（cpca，用于识别简体中文字符串中省，市和区）
import requests  # 导入网页分析的第三方模块
import datetime  # 导入获取今日日期的内置模块
import chinese_calendar  # 导入中国日历的第三方模块
import pandas as pd  # 导入pandas，用dataframe.to_excel向myLife.xlsx添加数据
from openpyxl import load_workbook  # 读写Excel引擎，没有用xlwings
from pyquery import PyQuery as pq
from bs4 import BeautifulSoup


def City2Id():
    # CityIdTbl：字典型，用来保存城市名称和对应的城市编号，如'北京': '101010100'
    # 可通过http://www.weather.com.cn/textFC/hb.shtml 查看当前天气，源码可查看到城市编号
    base_url = "http://www.weather.com.cn/textFC/{}.shtml"
    # sites 各个区简称，如hb代表华北，拼接到base_url形成完成链接，可访问
    sites = ["hb", "db", "hd", "hz", "hn", "xb", "xn", "gat"]

    CityIdTbl = {}
    for i in sites:
        # 读取url，并用utf-8编码以兼容中文
        # doc = pq(base_url.format(i), encoding="utf-8")
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"
        }
        doc = pq(url=base_url.format(i), headers=headers, encoding="utf-8")
        # 解析class="conMidtab"节点下的a节点并调用items方法以实现遍历，示例如下
        # <div class="conMidtab">
        # <a href="http://www.weather.com.cn/weather/101050101.shtml" target="_blank">哈尔滨</a></td>
        list0 = doc(".conMidtab a").items()
        for lst in list0:
            # 条件1：lst.text() in ( '返回顶部','详情') 剔除返回顶部','详情'类型
            # 条件2：lst.text() in CityIdTbl.keys() 避免重复的城市再次计入，setdefault也可刷新
            # 条件3：len(str(lst.attr('href')))< 49 剔除链接中无ID的场景
            # 如<a href="/textFC/hongkong.shtml" target="_blank">香港</a>
            if (
                lst.text() in ("返回顶部", "详情", CityIdTbl.keys(), "")
                or len(str(lst.attr("href"))) < 49
            ):
                # or len(str(lst.attr('href')))< 49 :
                continue
            else:
                city_name = lst.text()
                city_no = lst.attr("href")
                CityIdTbl.setdefault(city_name, city_no[34:43])
    return CityIdTbl


# 主程序
td = datetime.date.today()  # 获取今天日期
week_list = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]

# 获取本上网终端的城市名CityName
url = "https://ip.tool.chinaz.com/"  # 该网站可以查得联网设备IP地址及所在位置
response = requests.get(url)
response.encoding = response.apparent_encoding
content = response.text
# print(content)
addr = content[content.find("IP的物理位置") : content.find("ip138提供")]  # 大致筛选出归属地所在的字符串
# print(addr)
a = addr.find("infoLocation") + 14
b = addr.find("</em>")
addr = addr[a:b]
print("地址：", addr)
# 从地址中分离出城市名
location_str = []
location_str.append(addr)
df = cpca.transform(location_str)  # 该方法可以输入任意的可迭代类型（如list，pandas的Series类型等）
CityName = df.iat[0, 1][:-1]
print("你所在的城市：", CityName)

# 由城市名获取天气网所用到的城市代码
CityIdTbl = City2Id()
cityID = CityIdTbl.get(CityName)
url = "http://www.weather.com.cn/weather/{}.shtml"  # 中国天气网址
url = url.format(cityID)  # 用到cityID
print(CityName, cityID, url)

# 基于城市代码从中国天气网获取天气信息
htmsrc = requests.get(url, timeout=15)
htmsrc.encoding = htmsrc.apparent_encoding
myhtm = htmsrc.text
mysoup = BeautifulSoup(myhtm, "html.parser")
mydata = mysoup.body.find("div", {"id": "7d"}).find("ul")
day_list = mydata.find_all("li")
weather = day_list[0].find_all("p")

df = pd.DataFrame()  # 创建一个空的DataFrame
# 向空DataFrame中添加内容
df["日期"] = [td.strftime("%Y-%m-%d")]
df["星期"] = week_list[td.weekday()]
df["城市"] = CityName
df["地点"] = "校园"  # 这个程序不能自动获取所在城市的县区名，这里给了一个缺省值，必要时要改成真实地点
df["天气"] = weather[0].string
if weather[1].find("span") is None:
    df["最高温度"] = None
else:
    df["最高温度"] = weather[1].find("span").string
df["最低温度"] = weather[1].find("i").string[:-1]
df["是否工作日"] = chinese_calendar.is_workday(td)

print(df)

df1 = pd.read_excel(
    "myLife.xlsx", sheet_name="Scene", engine="openpyxl"
)  # 读取Scene工作表数据
Lastrow = df1.shape[0]  # 获取Scene工作表的总行数,接着在下一行Lastrow+1写入今天的场景数据。

# 建立写入对象,数据的写入
with pd.ExcelWriter(
    "myLife.xlsx", engine="openpyxl", mode="a", if_sheet_exists="overlay"
) as writer:
    df.to_excel(
        writer, sheet_name="Scene", startrow=Lastrow + 1, index=False, header=False
    )
