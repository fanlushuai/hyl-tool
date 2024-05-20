# 读取文件，校验格式。拖动进去来处理。
# 加上统计，原样输出，查看。
# 输出新文件
# 直接打印

from decimal import Decimal
import time
import os

# 文件夹目录
path = r"D:\360安全浏览器下载"

print("即将读取 " + path + " 文件夹中的最新文件")
time.sleep(4)

# 获取文件夹中所有的文件(名)，以列表形式返货
lists = os.listdir(path)
# print("未经处理的文件夹列表：\n %s \n" % lists)

# 按照key的关键字进行生序排列，lambda入参x作为lists列表的元素，获取文件最后的修改日期，
# 最后对lists以文件时间从小到大排序
lists.sort(key=lambda x: os.path.getmtime((path + "\\" + x)))

# 获取最新文件的绝对路径，列表中最后一个值,文件夹+文件名
file_new = os.path.join(path, lists[-1])
# print("时间排序后的的文件夹列表：\n %s \n" % lists)

print("最新文件路径:\n%s" % file_new)


import xlwings as xw

app = xw.App(visible=True, add_book=False)  # 界面设置
app.display_alerts = True  # 关闭提示信息
app.screen_updating = True  # 关闭显示更新

wb = app.books.open(file_new)
ws = wb.sheets[0]  # 0是第一个sheet
cell = ws.used_range.last_cell
rows = cell.row
columns = cell.column

# 获取表头，进行判断
x = ws.range((1, 1), (1, columns))

for v in x:
    print(v.value)
    if v.value == "收货人":
        senderColumns = v
    if v.value == "收货地址":
        addressColumns = v
    if v.value == "总件数":
        orderCountColumns = v
    if v.value == "件数" or v.value == "本车件数":
        countColumns = v
    if v.value == "体积" or v.value == "立方":
        spaceColumns = v

    # 判断是否是已清点，如果是，就退出
    if v.value == "已清点":
        print("已经处理过一次了,请重新导出文件")
        time.sleep(8)
        # wb.close()
        exit()

# 列，进行计算了。
# 以收货人为基准，计算收货人下面的所有数据
y = ws.range((1, senderColumns.column), (rows - 1, senderColumns.column))


# dic=[{k:name,v:[,,]}]
dic = []

sumSpaceForGaoPing = Decimal("0")
sumSpaceForChangZhi = Decimal("0")
sumSpaceForJinCheng = Decimal("0")
sumSpaceForYangCheng = Decimal("0")

for v in y:

    if v.value == "收货人":
        continue

    count = Decimal(str(ws[v.row - 1, countColumns.column - 1].value))

    if ws[v.row - 1, spaceColumns.column - 1].value != None:
        space = Decimal(str(ws[v.row - 1, spaceColumns.column - 1].value))
    else:
        space = Decimal("0")

    # print("件数%s 体积%s", v.value, count, space)

    address = str(ws[v.row - 1, addressColumns.column - 1].value)
    if address != None and address != "None":

        if "高平" in address:
            address = "高平"
            sumSpaceForGaoPing += space
        elif "长治" in address:
            address = "长治"
            sumSpaceForChangZhi += space
        elif "阳城" in address:
            address = "阳城"
            sumSpaceForYangCheng += space
        else:
            # 山西，默认都是空
            address = ""
            sumSpaceForJinCheng += space
    else:
        address = ""
        sumSpaceForJinCheng += space

    # 如果没有，就添加

    orderCount = Decimal(str(ws[v.row - 1, orderCountColumns.column - 1].value))

    if orderCount != count:
        print("件数不匹配", v.value, count, orderCount)
        r = ws.range((v.row, 1), (v.row, addressColumns.column))
        # r.api.Borders.LineStyle = 2
        # r.api.Borders.Weight = 3
        r.api.Font.Bold = True
        r.api.Font.Size = r.api.Font.Size - 1
        r.api.Characters.Font.Underline = 2

    has = 0
    for d in dic:
        if d["name"] == v.value:
            print("存在追加", v.value, count, space)
            d["orderCount"] = d["orderCount"] + 1
            d["count"] = d["count"] + count
            d["space"] = d["space"] + space
            has = 1
            break

    if has != 1:
        print("不存在新建", v.value, count, space)
        dic.append(
            {
                "name": v.value,
                "count": count,
                "space": space,
                "address": address,
                "orderCount": 1,
            }
        )

    # 否则，就累加
    # {"name": v.value, "count": count, "space": space}

print(dic)


# 排序一下。以体积排序
def sort_criteria(item):
    # 自定义排序函数的规则
    return item["space"]  # 返回列表中每个元素的第二个元素作为排序依据


dic = sorted(dic, key=sort_criteria, reverse=True)

ws.range(1, columns + 2).value = [
    "收货人",
    "总方数",
    "地址",
    "单数",
    "总件数",
    "已清点",
]
i = 2
sumCount = 0
sumSpace = 0
sumOrderCount = 0
for d in dic:
    ws.range(i, columns + 2).value = [
        d["name"],
        str(d["space"]),
        d["address"],
        str(d["orderCount"]),
        str(d["count"]),
        "",
    ]
    sumCount += d["count"]
    sumSpace += d["space"]
    sumOrderCount += d["orderCount"]
    i += 1
    # ws.range(2, columns + 5).value = d["name"]
    # ws.range(2, columns + 6).value = str(d["count"])
    # ws.range(2, columns + 7).value = str(d["space"])

ws.range(i, columns + 2).value = [
    ("共" + str(i - 2) + "行"),
    str(sumSpace),
    "",
    str(sumOrderCount),
    str(sumCount),
    "",
]

# 加边框
r = ws.range((1, columns + 2), (i, columns + 2 + 5))
r.api.Borders.LineStyle = 1
r.api.Borders.Weight = 2

fromLine = i + 2
toLine = i + 2

ws.range(i + 2, columns + 2).value = ["地址", "总方数"]
if sumSpaceForGaoPing > 0:
    toLine += 1
    fromLine += 1
    ws.range(fromLine, columns + 2).value = ["高平", str(sumSpaceForGaoPing)]
if sumSpaceForChangZhi > 0:
    toLine += 1
    fromLine += 1
    ws.range(fromLine, columns + 2).value = ["长治", str(sumSpaceForChangZhi)]
if sumSpaceForYangCheng > 0:
    toLine += 1
    fromLine += 1
    ws.range(fromLine, columns + 2).value = ["阳城", str(sumSpaceForYangCheng)]
if sumSpaceForJinCheng > 0:
    toLine += 1
    fromLine += 1
    ws.range(fromLine, columns + 2).value = ["晋城", str(sumSpaceForJinCheng)]

r = ws.range((i + 2, columns + 2), (toLine, columns + 2 + 1))
r.api.Borders.LineStyle = 1
r.api.Borders.Weight = 2

ws.autofit()

wb.save()

# 打印设置.横向打印即可

# wb.api.PrintOut()

# wb.close()
# app.quit()
# app.kill()

# auto-py-to-exe  auto-py-to-exe 运行
# https://pypi.org/project/auto-py-to-exe/

# Auto-Py-to-Exe完美打包python程序
# https://zhuanlan.zhihu.com/p/130328237
