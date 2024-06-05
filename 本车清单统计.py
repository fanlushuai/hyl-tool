import sqlite3


def createTable():
    # 创建连接
    conn = sqlite3.connect("doOrder.db")

    # 游标
    c = conn.cursor()
    try:
        # 如果存在，就会跳过
        # 建表语句
        c.execute(
            """CREATE TABLE orderDoHistory (
                    uniqueID TEXT,
                    orderID TEXT,
                    allCount INTEGER,
                    currentTimesCount INTEGER
            )"""
        )
        print("创建成功")
    except Exception as e:
        print(e)
    finally:
        conn.commit()
        conn.close()


def addHistory(uniqueID, orderID, allCount, currentTimesCount):
    conn = sqlite3.connect("doOrder.db")
    c = conn.cursor()
    # 插入语句
    sql = f"""INSERT INTO orderDoHistory (uniqueID, orderID, allCount,currentTimesCount) VALUES ('{uniqueID}', '{orderID}', {allCount},{currentTimesCount})"""
    try:
        c.execute(sql)
        print("添加成功")
    except Exception as e:
        print(e)
    finally:
        conn.commit()
        conn.close()


def getHistory(uniqueID, orderID):
    conn = sqlite3.connect("doOrder.db")
    c = conn.cursor()
    sql = f"""SELECT * FROM orderDoHistory WHERE uniqueID='{uniqueID}' AND orderID='{orderID}'"""
    try:
        result = c.execute(sql)
        return result.fetchall()[0]
    except Exception as e:
        print(e)
    finally:
        conn.close()


def updateHistory(uniqueID, orderID, currentTimesCount):
    conn = sqlite3.connect("doOrder.db")
    c = conn.cursor()
    sql = f"""UPDATE orderDoHistory SET currentTimesCount={currentTimesCount} WHERE uniqueID='{uniqueID}' AND orderID='{orderID}'"""
    try:
        c.execute(sql)
        print("更新成功")
    except Exception as e:
        print(e)
    finally:
        conn.commit()
        conn.close()


def getHistoryAllCurrentTimesCount(orderID):
    conn = sqlite3.connect("doOrder.db")
    c = conn.cursor()
    sql = f"""SELECT sum(currentTimesCount) FROM orderDoHistory WHERE orderID='{orderID}'"""
    try:
        result = c.execute(sql)
        return result.fetchall()[0][0]
    except Exception as e:
        print(e)
    finally:
        conn.close()


def genUniqueID(orderNumStr):
    import hashlib

    result = hashlib.md5(orderNumStr.encode())
    return result.hexdigest()


def getHistoryAll(orderID):
    conn = sqlite3.connect("doOrder.db")
    c = conn.cursor()
    sql = f"""SELECT * FROM orderDoHistory WHERE orderID='{orderID}'"""
    try:
        result = c.execute(sql)
        return result.fetchall()
    except Exception as e:
        print(e)
    finally:
        conn.close()


def getHistoryAll():
    conn = sqlite3.connect("doOrder.db")
    c = conn.cursor()
    sql = f"""SELECT * FROM orderDoHistory """
    try:
        result = c.execute(sql)
        return result.fetchall()
    except Exception as e:
        print(e)
    finally:
        conn.close()


def orderAllDone(uniqueID, orderID, allCount, currentTimesCount):

    # 处理车次运单的数据。一个车次，只能有一个运单。
    res = getHistory(uniqueID, orderID)
    if res is None:
        addHistory(uniqueID, orderID, allCount, currentTimesCount)
    else:
        updateHistory(uniqueID, orderID, currentTimesCount)

    # 一个运单，可以有很多条，不同的车次
    allCurrentTimesCount = getHistoryAllCurrentTimesCount(orderID)
    return allCurrentTimesCount >= allCount


createTable()

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
time.sleep(2)

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

import sys
import xlwings as xw
from xlwings.utils import rgb_to_int
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
    if v.value == "运单号":
        orderNoColumns = v
    if v.value == "收货人":
        senderColumns = v
    if v.value == "收货地址" or v.value == "收货人地址":
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
        from tkinter import messagebox
        messagebox.showinfo('信息', '放弃处理。【以前已经处理过了】')
        sys.exit(0)


# 生成此车次的唯一标记

orderNoAll = ws.range((2, orderNoColumns.column), (rows, orderNoColumns.column))
orderNumsStr = ""
for v in orderNoAll:
    print(v.value)
    orderNumsStr += v.value
uniqueId = genUniqueID(orderNumsStr)
print("uniqueId", uniqueId)
# 获取所有的部分发货数据


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

    orderNum = str(ws[v.row - 1, orderNoColumns.column - 1].value)
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

        # 各种样式参考：https://blog.csdn.net/NANGE007/article/details/124934344
        if orderAllDone(uniqueId, orderNum, orderCount, count):
            print("此被拆单全部处理完毕")
            r.api.Font.Size = r.api.Font.Size - 1
            # r.api.Font.Italic = True
            r.api.Font.ColorIndex = 10
            r.color = 169, 208, 142
            r.api.Font.Bold = False
        else:
            r.api.Font.Size = r.api.Font.Size - 1
            r.api.Font.ColorIndex = 3
            
            r.api.Font.Italic = True
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
