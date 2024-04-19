import matplotlib.pyplot as plt
# 收集 x 轴与 y 轴的数据
x = [1, 2, 3, 4, 5]
y = [2, 4, 6, 8, 10]
# 绘制折线图
plt.plot(x, y)
# 添加标题和坐标轴标签
plt.title('折线图示例')
plt.xlabel('X 轴')
plt.ylabel('Y 轴')
# 显示图形
plt.show()