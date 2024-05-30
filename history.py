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
# res = orderAllDone("fasdfsdfsdfdasdf", 100045, 10, 3)
# print(res)
