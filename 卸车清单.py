# 读取文件，校验格式。拖动进去来处理。
# 加上统计，原样输出，查看。
# 输出新文件
# 直接打印

from decimal import Decimal
import xlwings as xw

app = xw.App(visible=True, add_book=False)  # 界面设置
app.display_alerts = True  # 关闭提示信息
app.screen_updating = True  # 关闭显示更新

import os

# 文件夹目录
path = r"D:\360安全浏览器下载"

# 获取文件夹中所有的文件(名)，以列表形式返货
lists = os.listdir(path)
print("未经处理的文件夹列表：\n %s \n" % lists)

# 按照key的关键字进行生序排列，lambda入参x作为lists列表的元素，获取文件最后的修改日期，
# 最后对lists以文件时间从小到大排序
lists.sort(key=lambda x: os.path.getmtime((path + "\\" + x)))

# 获取最新文件的绝对路径，列表中最后一个值,文件夹+文件名
file_new = os.path.join(path, lists[-1])
print("时间排序后的的文件夹列表：\n %s \n" % lists)

print("最新文件路径:\n%s" % file_new)


wb = app.books.open(file_new)
ws = wb.sheets[0]  # 0是第一个sheet
cell = ws.used_range.last_cell
rows = cell.row
columns = cell.column

x = ws.range((1, 1), (1, columns))


for v in x:
    print(v.value)
    if v.value == "收货人":
        senderColumns = v
    if v.value == "件数":
        countColumns = v
    if v.value == "体积":
        spaceColumns = v
    if v.value == "总件数":
        print("已经处理过一次了")
        # wb.close()
        exit()

# 列，进行计算了。
y = ws.range((1, senderColumns.column), (rows - 1, senderColumns.column))


# dic=[{k:name,v:[,,]}]
dic = []

for v in y:

    if v.value == "收货人":
        continue

    count = Decimal(str(ws[v.row - 1, countColumns.column - 1].value))
    if ws[v.row - 1, spaceColumns.column - 1].value != None:

        space = Decimal(str(ws[v.row - 1, spaceColumns.column - 1].value))
    else:
        space = Decimal("0")

    # print("件数%s 体积%s", v.value, count, space)

    # 如果没有，就添加

    has = 0
    for d in dic:
        if d["name"] == v.value:
            print("存在追加", v.value, count, space)
            d["count"] = d["count"] + count
            d["space"] = d["space"] + space
            has = 1
            break

    if has != 1:
        print("不存在新建", v.value, count, space)

        dic.append({"name": v.value, "count": count, "space": space})

    # 否则，就累加
    # {"name": v.value, "count": count, "space": space}

print(dic)


# 排序一下。以体积排序
def sort_criteria(item):
    # 自定义排序函数的规则
    return item["space"]  # 返回列表中每个元素的第二个元素作为排序依据


dic = sorted(dic, key=sort_criteria, reverse=True)

ws.range(1, columns + 2).value = ["收货人", "总件数", "总方数", "已清点"]
i = 2
sumCount = 0
sumSpace = 0
for d in dic:
    ws.range(i, columns + 2).value = [d["name"], str(d["count"]), str(d["space"]), ""]
    sumCount += d["count"]
    sumSpace += d["space"]
    i += 1
    # ws.range(2, columns + 5).value = d["name"]
    # ws.range(2, columns + 6).value = str(d["count"])
    # ws.range(2, columns + 7).value = str(d["space"])

ws.range(i, columns + 2).value = [
    ("共" + str(i - 2) + "行"),
    str(sumCount),
    str(sumSpace),
    "",
]

# 加边框
r = ws.range((1, columns + 2), (i, columns + 2 + 2 + 1))
r.api.Borders.LineStyle = 1
r.api.Borders.Weight = 2

ws.autofit()

wb.save()

# 打印设置.横向打印即可

# wb.api.PrintOut()

# wb.close()
# app.quit()
# app.kill()
